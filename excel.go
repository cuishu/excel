package excel

import (
	"bytes"
	"errors"
	"reflect"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type Excel struct {
	Filename string
}

func getFieldName(field reflect.StructField) string {
	tag := field.Tag.Get("xlsx")
	if tag != "" {
		return tag
	}
	return field.Name
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
	f, err := excelize.OpenFile(e.Filename)
	if err != nil {
		return err
	}
	for i := 0; i < rt.NumField(); i++ {
		Sheet{Sheet: getFieldName(rt.Field(i))}.scanSheet(f, rv.Field(i).Addr())
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
