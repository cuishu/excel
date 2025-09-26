package excel

import (
	"bytes"
	"fmt"
	"io"
	"reflect"

	"github.com/cuishu/functools"
	excelize "github.com/xuri/excelize/v2"
)

func (s *Sheet) streamExportTitle(writer *excelize.StreamWriter, schema Schema, t reflect.Type) error {
	title := titleRow(schema, t)
	s.colCnt = len(title)
	if s.useTextStyle {
		writer.SetColStyle(1, s.colCnt, s.style)
	}
	return writer.SetRow("A1", functools.Map(func(v string) any {
		return &excelize.Cell{
			StyleID: s.style,
			Formula: "",
			Value:   v,
		}
	}, title))
}

func (s *Sheet) streamExportStruct(field reflect.Value) (any, error) {
	if field.Type() == picReflectType {
		panic("pic type not support")
	} else if field.Type() == cellReflectType {
		panic("cell type not support")
	}
	fun, ok := field.Type().MethodByName("MarshalXLSX")
	if !ok {
		fun, ok = field.Type().MethodByName("MarshalText")
	}
	if ok {
		res := fun.Func.Call([]reflect.Value{field})
		if res[1].Interface() != nil {
			err, ok := res[1].Interface().(error)
			if !ok {
				return nil, fmt.Errorf("%s has invalid return type", fun.Name)
			}
			return nil, err
		}
		return toString(res[0].Interface()), nil
	} else if isTime(field.Type()) {
		return field.Interface(), nil
	} else {
		panic("struct type must implement MarshalXLSX or MarshalText")
	}
}

func (s *Sheet) streamExportRow(writer *excelize.StreamWriter, obj reflect.Value, col column) error {
	var rowData []any = make([]any, 0, obj.NumField())
	t := obj.Type()
	for i := 0; i < obj.NumField(); i++ {
		field := obj.Field(i)
		if field.Kind() == reflect.Struct {
			if data, err := s.streamExportStruct(field); err != nil {
				return err
			} else {
				rowData = append(rowData, data)
			}
		} else {
			tag := getFieldName(t.Field(i))
			show, ok := s.filter[tag]
			if (len(s.filter) == 0) || (show && ok) {
				if field.NumMethod() > 0 {
					fun, ok := field.Type().MethodByName("MarshalXLSX")
					if !ok {
						fun, ok = field.Type().MethodByName("MarshalText")
					}
					if ok {
						res := fun.Func.Call([]reflect.Value{field})
						if res[1].Interface() != nil {
							err, ok := res[1].Interface().(error)
							if !ok {
								return fmt.Errorf("%s has invalid return type", fun.Name)
							}
							return err
						}
						rowData = append(rowData, toString(res[0].Interface()))
						continue
					}
				}
				rowData = append(rowData, toString(field.Interface()))
			}
		}
	}
	if err := writer.SetRow(col(), functools.Map(func(v any) any {
		return &excelize.Cell{
			StyleID: s.style,
			Formula: "",
			Value:   v,
		}
	}, rowData)); err != nil {
		return err
	}
	return nil
}

func (s *Sheet) streamExportRows(writer *excelize.StreamWriter, slice reflect.Value) error {
	rowNum := 1
	n := slice.Len()
	for i := range n {
		rowNum++
		obj := slice.Index(i)
		if err := s.streamExportRow(writer, obj, cellGenerator(rowNum)); err != nil {
			return err
		}
	}
	s.rowCnt = n
	return nil
}

func (s *Sheet) sheetStreamExport(f *excelize.File, rv reflect.Value) error {
	t := rv.Type().Elem().Elem()
	index, err := f.NewSheet(s.sheet)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)
	writer, err := f.NewStreamWriter(s.sheet)
	if err != nil {
		return err
	}

	if err := s.streamExportTitle(writer, s.filter, t); err != nil {
		return err
	}

	slice := rv.Elem()

	if err := s.streamExportRows(writer, slice); err != nil {
		return err
	}
	return writer.Flush()
}

func (s *Sheet) StreamExport(v any) (*bytes.Buffer, error) {
	f := excelize.NewFile()
	defer f.Close()
	rv := reflect.ValueOf(v)
	if err := s.sheetStreamExport(f, rv); err != nil {
		return nil, err
	}

	return f.WriteToBuffer()
}

func (s *Sheet) StreamExportTo(writer io.Writer, v any) error {
	f := excelize.NewFile()
	defer f.Close()
	rv := reflect.ValueOf(v)
	if err := s.sheetStreamExport(f, rv); err != nil {
		return err
	}
	_, err := f.WriteTo(writer)
	return err
}
