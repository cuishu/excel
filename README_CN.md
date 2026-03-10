# Excel

Excel 是一个用于轻松读写 Excel 文件的 Go 语言库。它提供了简单、基于结构体的 API，可以将电子表格数据直接映射到 Go 对象，反之亦然。底层基于高性能的 [excelize](https://github.com/qax-os/excelize) 库构建。

[![Go Reference](https://pkg.go.dev/badge/github.com/cuishu/excel.svg)](https://pkg.go.dev/github.com/cuishu/excel)
[![Go Report Card](https://goreportcard.com/badge/github.com/cuishu/excel)](https://goreportcard.com/report/github.com/cuishu/excel)

## 功能特性

- 将 Excel 工作表读取为结构体切片
- 将结构体切片写入 Excel 文件
- 支持单个文件中的多个工作表
- 通过结构体标签（`xlsx:"列名"`）自定义列映射
- 支持表头不在第一行时使用偏移量（offset）
- 支持通过实现 `MarshalXLSX` / `UnmarshalXLSX` 接口来自定义类型转换
- 针对大数据集的流式导出（速度更快，仅支持字符串值）
- 支持文件或 `io.Reader` / `io.Writer`

## 安装

```bash
go get github.com/cuishu/excel
```

## 快速开始

定义一个结构体，并用 `xlsx` 标签标记字段，标签值需与 Excel 文件中的列标题一致：

```go
type Human struct {
    ID   int    `xlsx:"id"`
    Name string `xlsx:"name"`
}
```

假设有一个 Excel 文件 `a.xlsx`，其中包含一个名为 `Sheet1` 的工作表：

| id | name   |
|----|--------|
| 1  | Smith  |

### 读取工作表

```go
var humans []Human

// 从文件读取
err := excel.Sheet{Filename: "a.xlsx", Sheet: "Sheet1"}.Scan(&humans)
if err != nil {
    // 处理错误
}

// 从 io.Reader 读取
reader := ... // 例如 http.Request 的 Body 或 os.File
err = excel.NewSheetFromReader(reader, "Sheet1").Scan(&humans)

for _, h := range humans {
    fmt.Println(h.Name)
}
```

### 写入工作表

```go
humans := []Human{
    {ID: 1, Name: "Smith"},
    {ID: 2, Name: "Jack"},
    {ID: 3, Name: "James"},
}

buff, err := (&excel.Sheet{Sheet: "Sheet1"}).Export(&humans)
if err != nil {
    // 处理错误
}
err = ioutil.WriteFile("a.xlsx", buff.Bytes(), 0644)
```

## 处理多个工作表

当 Excel 文件包含多个工作表时，可以使用一个容器结构体来存放每个工作表对应的切片。工作表名称通过 `xlsx` 标签指定。

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

### 读取多个工作表

```go
var data Data

// 从文件读取
err := excel.Excel{Filename: "b.xlsx"}.Scan(&data)

// 从 io.Reader 读取
err = excel.NewExcelFromReader(reader).Scan(&data)
```

### 写入多个工作表

```go
data := Data{
    Humans:  []Human{{ID: 1, Name: "Smith"}},
    Animals: []Animal{{Species: "Cat", Legs: 4}},
}

buff, err := excel.Excel{}.Export(&data)
if err != nil {
    // 处理错误
}
ioutil.WriteFile("b.xlsx", buff.Bytes(), 0644)
```

## 高级用法

### 表头偏移

如果列标题不在第一行，可以使用 `Offset(n)` 跳过指定行数：

```go
var items []MyStruct
err := excel.Sheet{Filename: "data.xlsx", Sheet: "Sheet1"}.
    Offset(2). // 跳过前两行
    Scan(&items)
```

### 自定义类型序列化

实现 `MarshalXLSX` 和 `UnmarshalXLSX` 接口可以控制自定义类型在 Excel 单元格中的读写方式。

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

之后就可以直接在结构体中使用 `Sex` 类型：

```go
type Person struct {
    Name string `xlsx:"name"`
    Sex  Sex    `xlsx:"sex"`
}
```

### 流式导出大数据

`StreamExport` 逐行写入数据，大幅降低内存占用。它速度更快，但要求所有值都能直接转换为字符串（不支持自定义序列化）。

```go
bigData := make([]Human, 1000000) // 大量数据

buff, err := (&excel.Sheet{Sheet: "Sheet1"}).StreamExport(&bigData)
if err != nil {
    // 处理错误
}
ioutil.WriteFile("large.xlsx", buff.Bytes(), 0644)
```

## 支持的数据类型

以下 Go 类型开箱即用：

- 有符号整数：`int`、`int8`、`int16`、`int32`、`int64`
- 无符号整数：`uint`、`uint8`、`uint16`、`uint32`、`uint64`
- 浮点数：`float32`、`float64`
- `string`
- `bool`
- `time.Time`（转换为 Excel 日期时间格式）

自定义类型可以通过实现上述序列化接口来支持。

## 许可证

MIT 许可证。详情请参阅 [LICENSE](LICENSE) 文件。

---

*Powered by [excelize](https://github.com/qax-os/excelize)*
