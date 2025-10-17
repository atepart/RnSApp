@echo off
setlocal

if "%1"=="" (
    exit /b
) else (
    set "Tag=%1"
)


rem Write version file
(
  echo __all__ = ["__version__", "REPO_SLUG"]
  echo __version__ = "%Tag%"
  echo REPO_SLUG = "atepart/RnSApp"
) > application\version.py

git add application\version.py
git commit -m "Release %Tag%"
git tag -a %Tag% -m "Version %Tag%"
git push origin --tags
git push

endlocal
