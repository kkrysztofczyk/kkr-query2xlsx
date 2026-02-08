# Release checklist

- Bump `APP_VERSION` in `main.pyw`
- Commit + push
- Tag: `git tag -a vX.Y.Z -m "Release vX.Y.Z"`
- Push tag: `git push origin vX.Y.Z`
- Check GitHub Actions
- Check Release assets (zip)
- Smoke test on Windows: unzip -> launch exe

ex. 
git tag -a v0.3.9 -m "v0.3.9"
git push origin v0.3.9