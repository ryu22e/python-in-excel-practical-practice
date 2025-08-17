# Python in Excelをより便利に使える実践プラクティスの紹介

## Requirements
* Git
* Python 3.10 or higher
* Your favorite editor

## How to Build
```bash
$ git clone {repository url}
$ cd {repository path}
$ python -m venv .venv
$ source .venv/bin/activate
(.venv)$ pip install -r requirements_dev.txt
(.venv)$ playwright install --with-deps chromium  # Run only once initially
(.venv)$ vim source/index.md  # Replace vim with your favorite editor
(.venv)$ make revealjs  # Build to build/revealjs/
```

If you want to automatically rebuild when changes are made, you can use `sphinx-autobuild`:

```bash
(.venv)$ sphinx-autobuild source build/revealjs -b revealjs  # Serve at http://localhost:8000
```
