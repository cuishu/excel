package excel

import (
	"bytes"
	"encoding"
	"io"
	"reflect"
	"time"

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

func (s *Sheet) streamExportStruct(field any) (any, error) {
	if _, ok := field.(Picture); ok {
		panic("pic type not support")
	} else if _, ok := field.(Cell); ok {
		panic("cell type not support")
	}
	if fun, ok := field.(XLSXMarshaler); ok {
		data, err := fun.MarshalXLSX()
		if err != nil {
			return nil, err
		}
		return data, nil
	} else if fun, ok := field.(encoding.TextMarshaler); ok {
		return fun.MarshalText()
	} else if t, ok := field.(time.Time); ok {
		return t, nil
	} else {
		panic("struct type must implement MarshalXLSX or MarshalText")
	}
}

func (s *Sheet) streamExportRow(writer *excelize.StreamWriter, obj reflect.Value, col column) error {
	var rowData []any = make([]any, 0, obj.NumField())
	t := obj.Type()
	for i := 0; i < obj.NumField(); i++ {
		field := obj.Field(i)
		fieldInterface := field.Interface()
		if field.Kind() == reflect.Struct {
			if data, err := s.streamExportStruct(fieldInterface); err != nil {
				return err
			} else {
				rowData = append(rowData, data)
			}
		} else {
			tag := getFieldName(t.Field(i))
			show, ok := s.filter[tag]
			if (len(s.filter) == 0) || (show && ok) {
				if fun, ok := fieldInterface.(XLSXMarshaler); ok {
					data, err := fun.MarshalXLSX()
					if err != nil {
						return err
					}
					rowData = append(rowData, string(data))
					continue
				} else if fun, ok := fieldInterface.(encoding.TextMarshaler); ok {
					data, err := fun.MarshalText()
					if err != nil {
						return err
					}
					rowData = append(rowData, string(data))
					continue
				}
				rowData = append(rowData, toString(fieldInterface))
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
