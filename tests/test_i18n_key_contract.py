import ast
import unittest
from pathlib import Path


class I18nKeyContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.repo_root = Path(__file__).resolve().parents[1]
        cls.main_path = cls.repo_root / "main.pyw"
        cls.source = cls.main_path.read_text(encoding="utf-8")
        cls.tree = ast.parse(cls.source, filename=str(cls.main_path))

    def _extract_i18n_en_keys(self) -> set[str]:
        for node in self.tree.body:
            if isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name) and node.target.id == "I18N":
                if node.value is None:
                    continue
                i18n_obj = ast.literal_eval(node.value)
                return set(i18n_obj["en"].keys())

            if isinstance(node, ast.Assign):
                if any(isinstance(target, ast.Name) and target.id == "I18N" for target in node.targets):
                    i18n_obj = ast.literal_eval(node.value)
                    return set(i18n_obj["en"].keys())

        self.fail("Unable to find literal I18N assignment in main.pyw")

    def _extract_literal_t_keys(self) -> set[str]:
        keys: set[str] = set()
        for node in ast.walk(self.tree):
            if not isinstance(node, ast.Call):
                continue
            if not isinstance(node.func, ast.Name) or node.func.id != "t":
                continue
            if not node.args:
                continue

            arg0 = node.args[0]
            if isinstance(arg0, ast.Constant) and isinstance(arg0.value, str):
                keys.add(arg0.value)
        return keys

    def test_all_literal_t_keys_exist_in_english_catalog(self):
        en_keys = self._extract_i18n_en_keys()
        used_keys = self._extract_literal_t_keys()

        missing = sorted(key for key in used_keys if key not in en_keys)
        self.assertEqual(
            missing,
            [],
            msg=(
                "Found t(\"...\") keys missing in I18N['en']: "
                + ", ".join(missing)
            ),
        )


if __name__ == "__main__":
    unittest.main()
