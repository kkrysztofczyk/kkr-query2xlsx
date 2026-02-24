import ast
import string
import unittest
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Optional


@dataclass(frozen=True)
class TCallSite:
    key: str
    lineno: int
    col: int
    line: str
    kw_names: tuple[str, ...]
    has_kwargs_unpack: bool


@dataclass(frozen=True)
class FormatStringIssue:
    lang: str
    key: str
    text: str
    error: str


def _read_main() -> tuple[Path, str, list[str]]:
    repo_root = Path(__file__).resolve().parents[1]
    main_path = repo_root / "main.pyw"
    src = main_path.read_text(encoding="utf-8")
    return main_path, src, src.splitlines()


def _extract_i18n_dict(tree: ast.Module) -> dict[str, dict[str, str]]:
    # I18N: ... = {...}  OR  I18N = {...}
    for node in tree.body:
        if isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            if node.target.id == "I18N" and node.value is not None:
                data = ast.literal_eval(node.value)
                return _validate_i18n(data)

        if isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Name) and tgt.id == "I18N":
                    data = ast.literal_eval(node.value)
                    return _validate_i18n(data)

    raise RuntimeError("Nie znaleziono literalnej definicji I18N w main.pyw.")


def _validate_i18n(data: Any) -> dict[str, dict[str, str]]:
    if not isinstance(data, dict):
        raise TypeError(f"I18N nie jest dict, tylko: {type(data)!r}")
    if "en" not in data:
        raise KeyError("I18N musi zawierać klucz 'en' (język referencyjny).")

    out: dict[str, dict[str, str]] = {}
    for lang, lang_dict in data.items():
        if not isinstance(lang, str):
            raise TypeError(f"I18N language code must be str, got: {type(lang)!r}")
        if not isinstance(lang_dict, dict):
            raise TypeError(f"I18N[{lang}] must be dict, got: {type(lang_dict)!r}")
        out[lang] = {}
        for k, v in lang_dict.items():
            if not isinstance(k, str) or not isinstance(v, str):
                raise TypeError(f"I18N[{lang}] must be dict[str,str]. Problem: {k!r} -> {type(v)!r}")
            out[lang][k] = v
    return out


def _collect_t_calls(tree: ast.AST, lines: list[str]) -> list[TCallSite]:
    out: list[TCallSite] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "t":
            continue
        if not node.args:
            continue

        key_node = node.args[0]
        key: Optional[str] = None
        if isinstance(key_node, ast.Constant) and isinstance(key_node.value, str):
            key = key_node.value
        elif isinstance(key_node, ast.Str):  # pragma: no cover (old Py)
            key = key_node.s

        if not key:
            continue

        lineno = int(getattr(node, "lineno", 0) or 0)
        col = int(getattr(node, "col_offset", 0) or 0)
        line = lines[lineno - 1].rstrip("\n") if 1 <= lineno <= len(lines) else ""

        kw_names: list[str] = []
        has_unpack = False
        for kw in (node.keywords or []):
            if kw.arg is None:
                has_unpack = True
            else:
                kw_names.append(kw.arg)

        out.append(
            TCallSite(
                key=key,
                lineno=lineno,
                col=col,
                line=line,
                kw_names=tuple(sorted(set(kw_names))),
                has_kwargs_unpack=has_unpack,
            )
        )
    return sorted(out, key=lambda c: (c.key, c.lineno, c.col))


def _placeholder_names(fmt: str) -> tuple[set[str], Optional[str]]:
    """
    Zwraca bazowe nazwy pól z format stringów:
    - {name} -> "name"
    - {user.name} -> "user"
    - {arr[0]} -> "arr"
    - {} -> "" (pozycyjny placeholder; ryzykowny przy .format(**kwargs))
    """
    formatter = string.Formatter()
    names: set[str] = set()
    try:
        for _lit, field_name, _spec, _conv in formatter.parse(fmt):
            if field_name is None:
                continue
            if field_name == "":
                names.add("")  # positional '{}'
                continue
            base = field_name.split(".", 1)[0].split("[", 1)[0]
            names.add(base)
    except ValueError as exc:
        return set(), str(exc)
    return names, None


def _short_list(xs: Iterable[str], limit: int = 10) -> str:
    xs = list(xs)
    if len(xs) <= limit:
        return ", ".join(xs)
    return ", ".join(xs[:limit]) + f" (+{len(xs) - limit})"


class I18nCheckTests(unittest.TestCase):
    def test_i18n_full_check(self):
        main_path, src, lines = _read_main()
        tree = ast.parse(src, filename=str(main_path))
        i18n = _extract_i18n_dict(tree)

        en = i18n["en"]
        en_keys = set(en.keys())

        # 1) Key completeness vs EN (dla wszystkich języków != en)
        key_errors: list[str] = []
        for lang, lang_dict in i18n.items():
            if lang == "en":
                continue
            lang_keys = set(lang_dict.keys())
            missing = sorted(en_keys - lang_keys)
            extra = sorted(lang_keys - en_keys)
            if missing:
                key_errors.append(f"[{lang}] missing keys: {_short_list(missing)}")
            if extra:
                key_errors.append(f"[{lang}] extra keys: {_short_list(extra)}")

        # 2) t("...") calls must exist in EN
        calls = _collect_t_calls(tree, lines)
        used_keys = sorted({c.key for c in calls})
        missing_used = sorted(set(used_keys) - en_keys)

        # 3) Placeholder compatibility EN vs each language + parse errors + positional '{}'
        fmt_issues: list[FormatStringIssue] = []
        ph_errors: list[str] = []
        callsite_errors: list[str] = []

        # cache placeholders per key per lang
        placeholders: dict[str, dict[str, set[str]]] = {lang: {} for lang in i18n.keys()}
        parse_err: dict[str, dict[str, Optional[str]]] = {lang: {} for lang in i18n.keys()}

        for lang, lang_dict in i18n.items():
            for k, text in lang_dict.items():
                ph, err = _placeholder_names(text)
                placeholders[lang][k] = ph
                parse_err[lang][k] = err
                if err:
                    fmt_issues.append(FormatStringIssue(lang=lang, key=k, text=text, error=err))

        # Compare placeholder sets vs EN for common keys
        for lang, lang_dict in i18n.items():
            if lang == "en":
                continue
            common = sorted(set(lang_dict.keys()) & en_keys)
            for k in common:
                en_ph = placeholders["en"].get(k, set())
                lang_ph = placeholders[lang].get(k, set())

                # '{}' w którymkolwiek języku = ryzyko IndexError przy format(**kwargs)
                if "" in en_ph or "" in lang_ph:
                    ph_errors.append(f"[{lang}] key={k} contains positional '{{}}' placeholder")

                if en_ph != lang_ph:
                    # pomijamy czyste różnice wynikające z błędu parsowania - i tak jest w fmt_issues
                    if parse_err["en"].get(k) or parse_err[lang].get(k):
                        continue
                    ph_errors.append(
                        f"[{lang}] key={k} placeholders mismatch: "
                        f"en={sorted(x for x in en_ph if x)} vs {lang}={sorted(x for x in lang_ph if x)}"
                    )

        # 4) Call-site safety: jeśli t("KEY", a=..., b=...) bez **kwargs,
        #    to EN placeholdery muszą być pokryte (bo t() robi .format(**kwargs) tylko gdy kwargs).
        en_placeholders_by_key = placeholders["en"]
        for c in calls:
            en_ph = en_placeholders_by_key.get(c.key)
            if not en_ph:
                continue

            # jeśli wywołanie bez kwargs i bez **kwargs -> t() NIE formatuje, więc nie ma KeyError
            if not c.kw_names and not c.has_kwargs_unpack:
                continue

            # **kwargs: nie da się statycznie sprawdzić
            if c.has_kwargs_unpack:
                continue

            missing_kwargs = sorted(x for x in en_ph if x and x not in set(c.kw_names))
            if missing_kwargs:
                callsite_errors.append(
                    f"t('{c.key}') at {c.lineno}:{c.col} missing kwargs={missing_kwargs} | {c.line.strip()}"
                )

        problems: list[str] = []
        problems.extend(key_errors)

        if missing_used:
            problems.append(f"t('...') keys missing in EN: {_short_list(missing_used)}")

        if fmt_issues:
            problems.append(
                "format string parse errors: "
                + _short_list([f"{x.lang}:{x.key}" for x in fmt_issues], limit=15)
            )

        problems.extend(ph_errors)

        if callsite_errors:
            problems.append("potential KeyError call sites:\n  - " + "\n  - ".join(callsite_errors[:20]))
            if len(callsite_errors) > 20:
                problems.append(f"(+{len(callsite_errors) - 20} more)")

        self.assertFalse(
            problems,
            msg="I18N check failed:\n- " + "\n- ".join(problems),
        )


if __name__ == "__main__":
    unittest.main()
