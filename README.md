# Excel

[中文](README_CN.md) | [English](README.md)

Excel is a Go library for reading and writing Excel files with ease. It provides a simple, struct-based API to map spreadsheet data directly to Go objects and vice versa. Built on top of the high-performance [excelize](https://github.com/qax-os/excelize) library.

[![Go Reference](https://pkg.go.dev/badge/github.com/cuishu/excel.svg)](https://pkg.go.dev/github.com/cuishu/excel)
[![Go Report Card](https://goreportcard.com/badge/github.com/cuishu/excel)](https://goreportcard.com/report/github.com/cuishu/excel)

## Features

- Read Excel sheets into slices of structs
- Write slices of structs to Excel files
- Support for multiple sheets in a single file
- Customizable column mapping via struct tags (`xlsx:"column_name"`)
- Handle headers not in the first row with offset
- Support for custom types by implementing `MarshalXLSX` / `UnmarshalXLSX`
- Stream export for large datasets (faster, string‑only values)
- Works with files or `io.Reader` / `io.Writer`

## Installation

```bash
go get github.com/cuishu/excel
```

## Quick Start

Define a struct with `xlsx` tags that match the column headers in your Excel file:

```go
type Human struct {
    ID   int    `xlsx:"id"`
    Name string `xlsx:"name"`
}
```

Assume an Excel file `a.xlsx` with a sheet named `Sheet1`:

| id | name   |
|----|--------|
| 1  | Smith  |

### Reading a Sheet

```go
var humans []Human

// From file
err := excel.NewSheetFromFile("a.xlsx", "Sheet1").Scan(&humans)
if err != nil {
    // handle error
}

// From io.Reader
reader := ... // e.g., http.Request body or os.File
err = excel.NewSheetFromReader(reader, "Sheet1").Scan(&humans)

for _, h := range humans {
    fmt.Println(h.Name)
}
```

### Writing a Sheet

```go
humans := []Human{
    {ID: 1, Name: "Smith"},
    {ID: 2, Name: "Jack"},
    {ID: 3, Name: "James"},
}

buff, err := excel.NewSheet("Sheet1").Export(&humans)
if err != nil {
    // handle error
}
err = ioutil.WriteFile("a.xlsx", buff.Bytes(), 0644)
```

## Working with Multiple Sheets

When your Excel file contains several sheets, use a container struct that holds slices for each sheet. The sheet name is specified in the `xlsx` tag.

```go
type Human struct {
    ID   int    `xlsx:"id"`
    Name string `xlsx:"name"`
}

type Animal struct {
    Species string `xlsx:"species"`
    Legs    int    `xlsx:"legs"`
}

type Data struct {
    Humans  []Human  `xlsx:"humans"`
    Animals []Animal `xlsx:"animals"`
}
```

### Reading Multiple Sheets

```go
var data Data

// From file
err := excel.NewExcel(Filename: "b.xlsx").Scan(&data)

// From io.Reader
err = excel.NewExcelFromReader(reader).Scan(&data)
```

### Writing Multiple Sheets

```go
data := Data{
    Humans:  []Human{{ID: 1, Name: "Smith"}},
    Animals: []Animal{{Species: "Cat", Legs: 4}},
}

buff, err := excel.Excel{}.Export(&data)
if err != nil {
    // handle error
}
ioutil.WriteFile("b.xlsx", buff.Bytes(), 0644)
```

## Advanced Usage

### Header Offset

If the column headers are not on the first row, use `Offset(n)` to skip rows:

```go
var items []MyStruct
err := excel.NewSheetFromFile("data.xlsx", "Sheet1").
    Offset(2). // skip first two rows
    Scan(&items)
```

### Custom Type Marshaling

Implement `MarshalXLSX` and `UnmarshalXLSX` to control how your custom types are converted to/from Excel cell values.

```go
type Sex int

const (
    Male Sex = iota + 1
    Female
)

func (s Sex) MarshalXLSX() ([]byte, error) {
    switch s {
    case Male:
        return []byte("Male"), nil
    case Female:
        return []byte("Female"), nil
    default:
        return nil, errors.New("invalid sex")
    }
}

func (s *Sex) UnmarshalXLSX(data []byte) error {
    switch string(data) {
    case "Male":
        *s = Male
    case "Female":
        *s = Female
    default:
        return errors.New("invalid sex")
    }
    return nil
}
```

Now you can use `Sex` directly in your struct:

```go
type Person struct {
    Name string `xlsx:"name"`
    Sex  Sex    `xlsx:"sex"`
}
```

### Stream Export for Large Data

`StreamExport` writes rows one by one, reducing memory usage. It is faster but only accepts values that can be directly converted to strings (no custom marshaling).

```go
bigData := make([]Human, 1000000) // large slice

buff, err := excel.NewSheet("Sheet1").StreamExport(&bigData)
if err != nil {
    // handle error
}
ioutil.WriteFile("large.xlsx", buff.Bytes(), 0644)
```

## Supported Data Types

The following Go types are supported out of the box:

- Signed integers: `int`, `int8`, `int16`, `int32`, `int64`
- Unsigned integers: `uint`, `uint8`, `uint16`, `uint32`, `uint64`
- Floating point: `float32`, `float64`
- `string`
- `bool`
- `time.Time` (converted to Excel date/time format)

Custom types can be supported by implementing the marshaling interfaces as shown above.

## License

MIT License. See [LICENSE](LICENSE) for details.

---

*Powered by [excelize](https://github.com/qax-os/excelize)*