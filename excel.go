package excel

import (
	"bytes"
	"errors"
	"io"
	"reflect"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

type Excel struct {
	Filename string
	reader   io.Reader
	offset   int
	style    int
}

func NewExcelFromReader(r io.Reader) *Excel {
	return &Excel{reader: r}
}

func (e *Excel) Offset(n int) *Excel {
	e.offset = n
	return e
}

func getFieldName(field reflect.StructField) string {
	tag := strings.TrimSpace(field.Tag.Get("xlsx"))
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

func (e Excel) Scan(v any) error {
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
	defer f.Close()
	for i := 0; i < rt.NumField(); i++ {
		Sheet{Sheet: getFieldName(rt.Field(i))}.Offset(e.offset).scanSheet(f, rv.Field(i).Addr())
	}

	return nil
}

func (e *Excel) export(f *excelize.File, v any) error {
	rv := reflect.ValueOf(v)

	if rv.Kind() != reflect.Ptr {
		return errors.New("param type must be ptr")
	}
	rv = rv.Elem()
	if rv.Kind() != reflect.Struct {
		return errors.New("param type must be struct ptr")
	}
	rt := rv.Type()
	deleteDefaultSheet := true
	for i := 0; i < rt.NumField(); i++ {
		sheet := &Sheet{Sheet: getFieldName(rt.Field(i))}
		if err := sheet.sheetExport(f, rv.Field(i).Addr()); err != nil {
			return err
		}
		if sheet.Sheet == defaultSheet {
			deleteDefaultSheet = false
		}
	}
	if deleteDefaultSheet {
		f.DeleteSheet(defaultSheet)
	}
	return nil
}

func (e Excel) Export(v any) (*bytes.Buffer, error) {
	f := excelize.NewFile()
	defer f.Close()
	if err := e.export(f, v); err != nil {
		return nil, err
	}
	return f.WriteToBuffer()
}

func (e Excel) ExportTo(w io.Writer, v any) error {
	f := excelize.NewFile()
	defer f.Close()
	if err := e.export(f, v); err != nil {
		return err
	}
	_, err := f.WriteTo(w)
	return err
}

func (e *Excel) streamExport(f *excelize.File, v any) error {
	rv := reflect.ValueOf(v)

	if rv.Kind() != reflect.Ptr {
		return errors.New("param type must be ptr")
	}
	rv = rv.Elem()
	if rv.Kind() != reflect.Struct {
		return errors.New("param type must be struct ptr")
	}
	rt := rv.Type()
	style, err := f.NewStyle(&excelize.Style{
		NumFmt: 49,
	})
	if err != nil {
		return err
	}
	e.style = style
	deleteDefaultSheet := true
	for i := 0; i < rt.NumField(); i++ {
		sheet := &Sheet{Sheet: getFieldName(rt.Field(i)), style: style}
		if err := sheet.sheetStreamExport(f, rv.Field(i).Addr()); err != nil {
			return err
		}
		if sheet.Sheet == defaultSheet {
			deleteDefaultSheet = false
		}
	}
	if deleteDefaultSheet {
		f.DeleteSheet(defaultSheet)
	}
	return nil
}

func (e Excel) StreamExport(v any) (*bytes.Buffer, error) {
	f := excelize.NewFile()
	defer f.Close()
	if err := e.streamExport(f, v); err != nil {
		return nil, err
	}
	return f.WriteToBuffer()
}

func (e Excel) StreamExportTo(w io.Writer, v any) error {
	f := excelize.NewFile()
	defer f.Close()
	if err := e.streamExport(f, v); err != nil {
		return err
	}
	_, err := f.WriteTo(w)
	return err
}
