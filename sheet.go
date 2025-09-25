package excel

import (
	"bytes"
	"errors"
	"fmt"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"reflect"
	"strconv"
	"strings"

	"github.com/cuishu/functools"
	excelize "github.com/xuri/excelize/v2"
)

const defaultSheet = "Sheet1"

type Schema map[string]bool

type Sheet struct {
	Filename      string
	Sheet         string
	filter        Schema
	offset        int
	reader        io.Reader
	style         int
	appendRowsCnt int
	rowCnt        int
	colCnt        int
}

func NewSheetFromReader(r io.Reader, sheet string) *Sheet {
	return &Sheet{Sheet: sheet, reader: r}
}

func (s *Sheet) AppendEmptyRows(n int) {
	s.appendRowsCnt = n
}

func (s Sheet) Offset(n int) Sheet {
	s.offset = n
	return s
}

func (s Sheet) excelizeOpen() (*excelize.File, error) {
	if s.Filename != "" {
		return excelize.OpenFile(s.Filename)
	} else if s.reader != nil {
		return excelize.OpenReader(s.reader)
	}
	return nil, errors.New("filename can not be empty")
}

func (s Sheet) scanSheet(f *excelize.File, rv reflect.Value) error {
	props, err := f.GetWorkbookProps()
	if err != nil {
		return err
	}

	date1904 := *props.Date1904

	t := rv.Type().Elem().Elem()

	rows, err := f.GetRows(s.Sheet)
	if err != nil {
		return err
	}
	var schema []string = make([]string, 0, t.NumField())
	var length int = len(rows)
	if length <= s.offset {
		return fmt.Errorf("file rows less than %d", s.offset+1)
	}
	rows = rows[s.offset:]
	length = len(rows)
	array := reflect.MakeSlice(rv.Type().Elem(), length-1, length)
	var indexArr []int = make([]int, 0, length)
	n := 0
	for i, row := range rows {
		var obj map[string]string = make(map[string]string)
		if i == 0 {
			schema = append(schema, functools.Map(func(s string) string { return strings.TrimSpace(s) }, row)...)
			continue
		}
		for j, cell := range row {
			value := strings.TrimSpace(cell)
			if value == "" || j >= len(schema) {
				continue
			}
			obj[schema[j]] = value
		}
		if len(obj) == 0 {
			continue
		}
		indexArr = append(indexArr, i-1)
		n++
		o := reflect.New(t)
		for j := 0; j < t.NumField(); j++ {
			tag := getFieldName(t.Field(j))
			valid := t.Field(j).Tag.Get("validate")
			field := o.Elem().Field(j).Addr().Interface()
			fieldType := reflect.TypeOf(field)
			fieldValue := reflect.ValueOf(field)
			value, ok := obj[tag]
			if !ok {
				continue
			}
			if fieldType.NumMethod() > 0 {
				if f, ok := fieldType.MethodByName("UnmarshalXLSX"); ok {
					in := reflect.New(f.Type.In(1)).Elem()
					in.SetBytes([]byte(value))
					values := f.Func.Call([]reflect.Value{fieldValue, in})
					if len(values) > 0 {
						err := values[0].Interface()
						if err != nil {
							return fmt.Errorf("line %d => %s: %s", i, err.(error).Error(), value)
						}
					}
					goto validate
				}
			}
			if rv, err := getReflectValue(value, fieldType.Elem()); err == nil {
				if isTime(fieldType.Elem()) {
					// styleID := s.timeStyle(f, rv)
					for col, elem := range row {
						if elem == value {
							cellName, err := excelize.CoordinatesToCellName(col+1, i+s.offset+1)
							if err != nil {
								return fmt.Errorf("line %d => %s: %s", i, err.Error(), value)
							}
							// f.SetCellStyle(s.Sheet, cellName, cellName, styleID)
							value, err := f.GetCellValue(s.Sheet, cellName, excelize.Options{RawCellValue: true})
							if err != nil {
								return fmt.Errorf("line %d => %s: %s", i, err.Error(), value)
							}
							v, err := strconv.ParseFloat(value, 64)
							if err != nil {
								return fmt.Errorf("line %d => %s: %s", i, err.Error(), value)
							}
							t, err := excelize.ExcelDateToTime(v, date1904)
							if err != nil {
								return fmt.Errorf("line %d => %s: %s", i, err.Error(), value)
							}
							o.Elem().Field(j).Set(reflect.ValueOf(t))
						}
					}
					goto validate
				}
				if fieldType.Elem() == picReflectType {
					var pictures []Picture
					var err error
					pics, err := f.GetPictures(s.Sheet, cell(i+1, j))
					if err != nil {
						return err
					}
					pictures = functools.Map(func(pic excelize.Picture) Picture {
						return Picture{
							File:     pic.File,
							Format:   (*PicFormat)(pic.Format),
							withPath: false,
						}
					}, pics)
					rv = reflect.ValueOf(pictures)
				}
				o.Elem().Field(j).Set(rv)
			} else {
				return fmt.Errorf("line %d => %s: %s", i, err.Error(), value)
			}
		validate:
			if valid != "" {
				if o.Elem().Field(j).Interface() != nil {
					if err := validate.Var(o.Elem().Field(j).Interface(), valid); err != nil {
						return fmt.Errorf("line %d => %s: %v %s", i, tag, o.Elem().Field(j).Interface(), err.Error())
					}
				}
			}
		}
		array.Index(i - 1).Set(o.Elem())
	}
	items := reflect.MakeSlice(rv.Type().Elem(), n, n)
	for i, index := range indexArr {
		items.Index(i).Set(array.Index(index))
	}
	rv.Elem().Set(items)

	return nil
}

func (s Sheet) Scan(v any) error {
	rv := reflect.ValueOf(v)
	if rv.Kind() != reflect.Ptr || rv.IsNil() || rv.Type().Elem().Kind() != reflect.Slice {
		panic("param must be slice pointer")
	}
	f, err := s.excelizeOpen()
	if err != nil {
		return err
	}
	defer f.Close()
	return s.scanSheet(f, rv)
}

type column func() string

func cell(x, y int) string {
	return fmt.Sprintf("%s%d", toTwentySix(y), x)
}

func cellGenerator(line int) column {
	i := 0
	return func() string {
		i++
		return cell(line, i)
	}
}

func titleRow(schema Schema, t reflect.Type) []string {
	var title []string = make([]string, 0, t.NumField())
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		tag := getFieldName(field)
		show, ok := schema[tag]
		if (len(schema) == 0) || (show && ok) {
			title = append(title, tag)
		}
	}
	return title
}

func (s *Sheet) exportTitle(f *excelize.File, schema Schema, sheet string, t reflect.Type, col column) {
	title := titleRow(schema, t)
	s.colCnt = len(title)
	for _, v := range title {
		colIdx := col()
		f.SetCellStyle(sheet, colIdx, colIdx, s.style)
		f.SetCellStr(sheet, colIdx, v)
	}
}

func (s *Sheet) exportPic(f *excelize.File, field reflect.Value, col column) error {
	pic := field.Interface().(Picture)
	if pic.withPath {
		if err := f.AddPicture(s.Sheet, col(), pic.Name, (*excelize.GraphicOptions)(pic.Format)); err != nil {
			return err
		}
	} else {
		if err := f.AddPictureFromBytes(s.Sheet, col(),
			&excelize.Picture{
				File:   pic.File,
				Format: (*excelize.GraphicOptions)(pic.Format),
			}); err != nil {
			return err
		}
	}
	return nil
}

func (s *Sheet) exportCell(f *excelize.File, field reflect.Value, col column) error {
	c := field.Interface().(Cell)
	if c.HyperLink.Link != "" {
		column := col()
		f.SetCellStr(s.Sheet, column, c.Value)
		f.SetCellHyperLink(s.Sheet, column, c.HyperLink.Link, string(c.HyperLink.Type))
		if c.Style != nil {
			style, err := f.NewStyle(c.Style)
			if err != nil {
				return err
			}
			if err := f.SetCellStyle(s.Sheet, column, column, style); err != nil {
				return err
			}
		}
	}
	return nil
}

func (s *Sheet) exportStruct(f *excelize.File, field reflect.Value, col column) error {
	if field.Type() == picReflectType {
		return s.exportPic(f, field, col)
	} else if field.Type() == cellReflectType {
		return s.exportCell(f, field, col)
	}
	if fun, ok := field.Type().MethodByName("MarshalXLSX"); ok {
		res := fun.Func.Call([]reflect.Value{field})
		if res[1].Interface() != nil {
			err, ok := res[1].Interface().(error)
			if !ok {
				return fmt.Errorf("%s has invalid return type", fun.Name)
			}
			return err
		}
		f.SetCellStr(s.Sheet, col(), toString(res[0].Interface()))
	} else if isTime(field.Type()) {
		f.SetCellValue(s.Sheet, col(), field.Interface())
	} else {
		panic("struct type must implement MarshalXLSX")
	}
	return nil
}

func (s *Sheet) exportRow(f *excelize.File, obj reflect.Value, col column) error {
	t := obj.Type()
	for i := 0; i < obj.NumField(); i++ {
		field := obj.Field(i)
		if field.Kind() == reflect.Struct {
			if err := s.exportStruct(f, field, col); err != nil {
				return err
			}
		} else {
			tag := getFieldName(t.Field(i))
			show, ok := s.filter[tag]
			if (len(s.filter) == 0) || (show && ok) {
				if field.NumMethod() > 0 {
					if fun, ok := field.Type().MethodByName("MarshalXLSX"); ok {
						res := fun.Func.Call([]reflect.Value{field})
						if res[1].Interface() != nil {
							err, ok := res[1].Interface().(error)
							if !ok {
								return fmt.Errorf("%s has invalid return type", fun.Name)
							}
							return err
						}
						colIdx := col()
						f.SetCellStyle(s.Sheet, colIdx, colIdx, s.style)
						f.SetCellStr(s.Sheet, colIdx, toString(res[0].Interface()))
						continue
					}
				}
				colIdx := col()
				f.SetCellStyle(s.Sheet, colIdx, colIdx, s.style)
				f.SetCellStr(s.Sheet, colIdx, toString(field.Interface()))
			}
		}
	}
	return nil
}

func (s *Sheet) exportRows(f *excelize.File, slice reflect.Value) error {
	rowNum := 1
	n := slice.Len()
	for i := 0; i < n; i++ {
		rowNum++
		obj := slice.Index(i)
		if err := s.exportRow(f, obj, cellGenerator(rowNum)); err != nil {
			return err
		}
	}
	s.rowCnt = n
	return nil
}

func (s *Sheet) sheetExport(f *excelize.File, rv reflect.Value) error {
	t := rv.Type().Elem().Elem()

	sheet, err := f.NewSheet(s.Sheet)
	if err != nil {
		return err
	}
	f.SetActiveSheet(sheet)

	s.exportTitle(f, s.filter, s.Sheet, t, cellGenerator(1))

	slice := rv.Elem()

	if err := s.exportRows(f, slice); err != nil {
		return err
	}

	return nil
}

func (s *Sheet) export(f *excelize.File, v any) error {
	if s.Sheet == "" {
		s.Sheet = defaultSheet
	}
	rv := reflect.ValueOf(v)
	if rv.Kind() != reflect.Pointer || rv.IsNil() || rv.Type().Elem().Kind() != reflect.Slice {
		panic("param must be slice ptr")
	}
	if err := s.sheetExport(f, rv); err != nil {
		return err
	}
	offset := 1 + s.offset + s.rowCnt
	n := s.appendRowsCnt + offset
	f.SetRowStyle(s.Sheet, offset, n, s.style)
	if s.Sheet != defaultSheet {
		f.DeleteSheet(defaultSheet)
	}
	return nil
}

func (s *Sheet) Export(v any) (*bytes.Buffer, error) {
	f := excelize.NewFile()
	defer f.Close()
	style, err := f.NewStyle(&excelize.Style{
		NumFmt: 49,
	})
	if err != nil {
		return nil, err
	}
	s.style = style
	if err := s.export(f, v); err != nil {
		return nil, err
	}
	return f.WriteToBuffer()
}

func (s *Sheet) ExportTo(w io.Writer, v any) error {
	f := excelize.NewFile()
	defer f.Close()
	style, err := f.NewStyle(&excelize.Style{
		NumFmt: 49,
	})
	if err != nil {
		return err
	}
	s.style = style
	if err := s.export(f, v); err != nil {
		return err
	}
	_, err = f.WriteTo(w)
	return err
}

func (s *Sheet) Filter(schema Schema) *Sheet {
	s.filter = schema
	return s
}
