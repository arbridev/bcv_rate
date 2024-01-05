# BCV rate

BCV rate is a utility that reports/outputs to a single Excel file the daily USD currency exchange rate to VES (Bolívares) from the Banco Central de Venezuela statistics, given their complete quarterly reports of a year.

## Dependency installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install dependencies.

```bash
pip install openpyxl
pip install python-dateutil
```

## Usage

Requires input files to be converted to `.xlsx` format and be available on the `input` directory. Original files are named in the format `2_1_2a23_smc.xls` (in this case `a23` refers to the quarter "a" of year "2023"). Converted required files will only change their extension to `.xlsx`.

Default arguments (in this case for the year `2023`, with files on the root `input` directory)
```bash
python main.py
```
Specified arguments (in this case for the year `2022`, with files on the `input/2022` directory)
```bash
python main.py -y 2022 -p 2022
```

## Acknowledgement

Banco Central de Venezuela for statistics publicly shared on their website (https://www.bcv.org.ve/estadisticas/tipo-cambio-de-referencia-smc).

## License

[MIT](https://choosealicense.com/licenses/mit/)