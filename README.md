# excel

Excel is a library to assist in manipulating Excel file.

Powered by github.com/qax-os/excelize

## Get excel

```bash
go get github.com/cuishu/excel
```

## How to use


The first line of the excel file must be the same with  the tag of Go struct

**For example**

We have a file named ```a.xlsx```

Sheet1
|id|name|
|:-:|:-:|
|1|Smith|


```go
type Human {
    ID   int    `xlsx:"id"`
    Name string `xlsx:"name"`
}
```

**Read Sheet**

```go
var humans []Human
// Read from file
Sheet{Filename: "a.xlsx", Sheet: "Sheet1"}.Scan(&humans)
// Read from reader
NewSheetFromReader(reader, "Sheet1").Scan(&humans)

for _, human := range humans {
  fmt.Println(human.Name)
  ...
}
```

**Go Slice to excel file**

```go
var humans []Human
humans = append(Human{ID: 1, Name: "Smith"})
humans = append(Human{ID: 2, Name: "Jack"})
humans = append(Human{ID: 3, Name: "James"})
    
buff, err := (&Sheet{Sheet: "Sheet1"}).Export(&users)

ioutil.WriteFile("a.xlsx", buff.Bytes(), 0644)
```

**Supported data types**

```go
int int8 int16 int32 int64
uint uint8 uint16 uint32 uint64
float32 float64
string
bool
```

## Read write the whole file

If the file has more than one sheet, we should use a struct receive them.

**Example**

```go
type Human struct {
	ID   int    `xlsx:"id"`
	Name string `xlsx:"name"`
}

type Animal Human

type Example struct {
	Humans  []Human  `xlsx:"humans"`
	Animals []Animal `xlsx:"animals"`
}

var example Example

// Read
(Excel{Filename: "b.xlsx"}).Scan(&example)
// Read from io.Reader
NewExcelFromReader(reader).Scan(&example)

// To bytes buffer
Excel{}.Export(&example)
```

## Offset
If the header is not in the first row, then you should use offset, which defaults to 0

```go
e := Sheet{Filename: "a.xlsx", Sheet: "Sheet3"}
var ss []TestObject
e.Offset(1).Scan(&ss)
...
```

## Advance

You can use custom types to implement MarshalXLSX and UnmarshalXLSX to implement type convert.

**For example**

```go
const (
	Male = 1
	Female = 2
)

type Sex int

func (sex Sex) MarshalXLSX() ([]byte, error) {
	if sex == Male {
		return []byte("Male"), nil
	} else if sex == Female {
		return []byte("Female"), nil
	}
	return nil, errors.New("invalid value")
}

func (sex *Sex) UnmarshalXLSX(data []byte) error {
	s := string(data)
	if s == "Male" {
		*sex = Male
		return nil
	}
	if s == "Female" {
		*sex = Female
		return nil
	}
	return errors.New("invalid value")
}
```