excel2xml
===
Convert `.xlsx` file to `.xml` file in a convenient Alpine Linux Docker image.

## Runnning
```bash
alias excel2xml='docker run --rm -it -v `pwd`:/pwd ubergarm/excel2xml'
excel2xml infile.xlsx outfile.xml
```

## Building
```bash
docker build -t ubergarm/excel2xml .
```

## Style Guide
```bash
flake8 --max-line-length=120 excel2xml.py
```

## References
* [pandas.read_excel](http://pandas.pydata.org/pandas-docs/stable/generated/pandas.read_excel.html)
