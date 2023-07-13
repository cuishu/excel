package excel

import (
	"bytes"
	"errors"
	"io"
	"reflect"

	excelize "github.com/xuri/excelize/v2"
)

type Excel struct {
	Filename string
	reader   io.Reader
	offset   int
}

func NewExcelFromReader(r io.Reader) *Excel {
	return &Excel{reader: r}
}

func (e *Excel) Offset(n int) *Excel {
	e.offset = n
	return e
}

func getFieldName(field reflect.StructField) string {
	tag := field.Tag.Get("xlsx")
	if tag != "" {
		return tag
	}
	return field.Name
}

func (e Excel) excelizeOpen() (*excelize.File, error) {
	if e.Filename != "" {
		return excelize.OpenFile(e.Filename)
	} else if e.reader != nil {
		return excelize.OpenReader(e.reader)
	}
	return nil, errors.New("filename can not be empty")
}

func (e Excel) Scan(v interface{}) error {
	rv := reflect.ValueOf(v)

	if rv.Kind() != reflect.Ptr {
		return errors.New("param type must be ptr")
	}
	rv = rv.Elem()
	if rv.Kind() != reflect.Struct {
		return errors.New("param type must be struct ptr")
	}
	rt := rv.Type()
	f, err := e.excelizeOpen()
	if err != nil {
		return err
	}
	for i := 0; i < rt.NumField(); i++ {
		Sheet{Sheet: getFieldName(rt.Field(i))}.Offset(e.offset).scanSheet(f, rv.Field(i).Addr())
	}

	return nil
}

func (e Excel) Export(v interface{}) (*bytes.Buffer, error) {
	rv := reflect.ValueOf(v)

	if rv.Kind() != reflect.Ptr {
		return nil, errors.New("param type must be ptr")
	}
	rv = rv.Elem()
	if rv.Kind() != reflect.Struct {
		return nil, errors.New("param type must be struct ptr")
	}
	rt := rv.Type()
	f := excelize.NewFile()
	deleteDefaultSheet := true
	for i := 0; i < rt.NumField(); i++ {
		sheet := &Sheet{Sheet: getFieldName(rt.Field(i))}
		if err := sheet.sheetExport(f, rv.Field(i).Addr()); err != nil {
			return nil, err
		}
		if sheet.Sheet == defaultSheet {
			deleteDefaultSheet = false
		}
	}
	if deleteDefaultSheet {
		f.DeleteSheet(defaultSheet)
	}
	return f.WriteToBuffer()
}
