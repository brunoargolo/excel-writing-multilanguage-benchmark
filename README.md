
# Writing Large xlsx excel files  

This repo tries to answer which language/lib combination is the fastest when writing very large xlsx files.

Here we explore some of the top excel libs from nodejs, go, rust and python feel free to add your own.

The results are listed below in absolute time measured by each script.

All tests were executed in a MacBook Pro, with an M3 chip and 36gb of ram.


## Details
The repo has 4 projects, each containing an equivalent script to generate a xlsx file based on the same dataset of roughly 1 million rows with 7 columns.

Each project is run in two modes:

**single-sheet**: 1 sheet with 1 million records is printed to a brand new xlsx file

**multi-sheet**: 9 sheets with 1 million records each are printed to the same brand new xlsx file

For multisheet approach all script try to write every sheet in parallel


## Results 
   
Language + Lib | 1 sheet | 9 sheets
--- | --- | --- 
Nodejs with exceljs | 5s | 41s
Go with excelize | 5s | 30s
Python3 with xlsxwriter | 26s | 236s
Rust with rust_xlsxwriter | 5s | 40s

## How to Run
Clone the repo. 
Making sure you have all the minimun features to run programs in python, nodejs, go and rust.

**Nodejs**
```
cd node_exceljs
npm ci
# single-sheet
node main.js
# multi-sheet
N_SHEETS=9 node main.js
```
**GO**
```
cd go_excelize
go mod init main.go
go get github.com/xuri/excelize/v2
# single-sheet
go run .
# multi-sheet
N_SHEETS=9 go run .
```
**Python**
```
cd python_xlsxwriter
python3 -m venv .venv
source .venv/bin/activate
# single-sheet
python main.py
# multi-sheet
N_SHEETS=9 python main.py
```
**Rust**
```
cd rust_xlsxwriter
# single-sheet
cargo run --release
# multi-sheet
N_SHEETS=9 cargo run --release
```