# Excel to JSON

Transform excel file to JSON translating and/or mapping keys.

## Usage

```bash
$ node src/index.js file.xls -d dict.json -m map_keys.json
```

- `file.xls` - Excel file
- `dict.json` - Translate key file
- `map_keys.json` - Map key file

### File example

Excel (excel.xlsx):

| User Name | Age | Company | Role | id |
| --------- | --- | ------- | ---- | --- |
| Bill Gates | 62 | Microsoft | director | 1 |

Dict (dict.json):
```json
{
  "User Name": "user_name",
  "Age": "age",
  "Company": "company",
  "Role": "role",
  "id": "id"
}
```

Map (map.json):
```json
{
  "User Name": "user.name",
  "Age": "user.age",
  "Company": "user.company.name",
  "Role": "user.company.role",
  "id": "id"
}
```

Runing: `$ node src/index.js excel.xlsx -d dict.json`
Result:
```json
 [
  {
    "user_name": "Bill Gates",
    "age": "62",
    "company": "Microsoft",
    "role": "director",
    "id": "1"
  }
]
```

Runing: `$ node src/index.js excel.xlsx -m map.json`
Result:
```json
[
  {
    "user": {
      "name": "Bill Gates",
      "age": "62",
      "company": {
        "name": "Microsoft",
        "role": "director"
      }
    },
    "id": "1"
  }
]
```

Runing `$ node src/index.js excel.xlsx -k`
Result (dict.json):
```json
{
  "User Name": "",
  "Age": "",
  "Company": "",
  "Role": "",
  "id": ""
}
```

## Options

-  **--createMap** or **-r**<br>To create a json file to map the keys
-  **--createDict** or **-k**<br>To create a json file to translate the keys
-  **--file** or **-f**<br>Excel file
-  **--dictFile** or **-d**<br>Json file with key dictionary.
-  **--mapFile** or **-m**<br>Json file with key map.
-  **--keysLine** or **-l**<br>Line number (excel) of keys.
-  **--help** or **-h**<br>Show this usage guide.