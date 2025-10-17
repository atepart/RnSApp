#! /bin/bash

while getopts "t:" arg; do
  case $arg in
    t) Tag=$OPTARG;;
  esac
done

# Write version into application/version.py
cat > application/version.py <<EOF
__all__ = ["__version__", "REPO_SLUG"]
__version__ = "$Tag"
REPO_SLUG = "atepart/RnSApp"
EOF

git add application/version.py
git commit -m "Release $Tag"
git tag -a $Tag -m "Version $Tag"

# Push
git push origin --tags
git push
