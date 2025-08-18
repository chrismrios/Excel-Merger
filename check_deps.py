import pkgutil, sys
reqs = ['streamlit','pandas','openpyxl','rapidfuzz','xlrd']
missing = [r for r in reqs if not pkgutil.find_loader(r)]
sys.exit(1 if missing else 0)