echo "Turning python into pdf file"
pyfile="$1"
set guifont=Courier
vim $pyfile.py -c ":hardcopy > $pyfile.ps" -c ":q"
ps2pdf $pyfile.ps $pyfile-pycode.pdf
echo "$pyfile-pycode.pdf made"
