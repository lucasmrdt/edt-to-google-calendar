# EDT to Google Calendar

## Installation

```bash
git clone https://github.com/lucasmrdt/edt-to-google-calendar
cd edt-to-google-calendar
pip install -r requirements.txt --user
```

## Help

```bash
./edt2google -h
```

## Example

|activity|group|
|:-:|:-:|
|PFA|Group 3|
|GLA|Group 1B|
|System|Group 2|
|Network|Group 1B|
|Logical|Group 1|
|Algorithm|Group 1B|



```bash
./edt2google --pfa g3 --gla g1 --sys g2 --net g1 --log g1 --algo g1  assets/edt-l3-info.xlsx 

```
