package excel

import (
	"bytes"
	"encoding"
	"errors"
	"fmt"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/cuishu/functools"
	excelize "github.com/xuri/excelize/v2"
)

const defaultSheet = "Sheet1"

type XLSXMarshaler interface {
	MarshalXLSX() ([]byte, error)
}

type XLSXUnmarshaler interface {
	UnmarshalXLSX(data []byte) error
}

type Schema map[string]bool

type Row struct {
	ID   int
	Data map[string]string
}

func (row Row) Get(key string) string {
	if v, ok := row.Data[key]; ok {
		return v
	}
	return ""
}

type Error struct {
	Row  Row
	mesg string
}

func (e Error) Error() string {
	return e.mesg
}

type Sheet struct {
	filename      string
	sheet         string
	title         []string
	errors        []Error
	filter        Schema
	offset        int
	reader        io.Reader
	style         int
	rowCnt        int
	colCnt        int
	useTextStyle  bool
	collectErrors bool
}

// NewSheet creates a new Sheet.
func NewSheet(sheet string) *Sheet {
	return &Sheet{sheet: sheet}
}

// NewSheetFromFile creates a new Sheet from a file.
func NewSheetFromFile(filename, sheet string) *Sheet {
	return &Sheet{sheet: sheet, filename: filename}
}

// NewSheetFromReader creates a new Sheet from a io.Reader.
func NewSheetFromReader(r io.Reader, sheet string) *Sheet {
	return &Sheet{sheet: sheet, reader: r}
}

// UseTextStyle sets the style of the cell to text.
func (s *Sheet) UseTextStyle() *Sheet {
	s.useTextStyle = true
	return s
}

// CollectErrors collects errors while scanning the sheet.
func (s *Sheet) CollectErrors() *Sheet {
	s.collectErrors = true
	return s
}

// Offset sets the offset of the sheet.
func (s *Sheet) Offset(n int) *Sheet {
	s.offset = n
	return s
}

// Errors returns the errors collected while scanning the sheet.
func (s *Sheet) Errors() []Error {
	return s.errors
}

func (s *Sheet) excelizeOpen() (*excelize.File, error) {
	if s.filename != "" {
		return excelize.OpenFile(s.filename)
	} else if s.reader != nil {
		return excelize.OpenReader(s.reader)
	}
	return nil, errors.New("filename can not be empty")
}

func (s *Sheet) scanTime(f *excelize.File, col int, i int, elem string, value string, date1904 bool) (time.Time, error) {
	if elem == value {
		cellName, err := excelize.CoordinatesToCellName(col, i+s.offset+1)
		if err != nil {
			return time.Time{}, err
		}
		// f.SetCellStyle(s.Sheet, cellName, cellName, styleID)
		value, err := f.GetCellValue(s.sheet, cellName, excelize.Options{RawCellValue: true})
		if err != nil {
			return time.Time{}, err
		}
		v, err := strconv.ParseFloat(value, 64)
		if err != nil {
			return time.Time{}, err
		}
		t, err := excelize.ExcelDateToTime(v, date1904)
		if err != nil {
			return time.Time{}, err
		}
		return t, nil
	}
	return time.Time{}, nil
}

func isDate1904(props excelize.WorkbookPropsOptions) bool {
	var date1904 bool
	if props.Date1904 != nil {
		date1904 = *props.Date1904
	}
	return date1904
}

type dataArray struct {
	array    []reflect.Value
	indexArr []int
}

func newDataArray(length int) dataArray {
	return dataArray{
		array:    make([]reflect.Value, length-1, length),
		indexArr: make([]int, 0, length),
	}
}

func createObjectMap(schema []string, row []string) map[string]string {
	obj := make(map[string]string)
	for j, cell := range row {
		value := strings.TrimSpace(cell)
		if j >= len(schema) {
			continue
		}
		obj[schema[j]] = value
	}
	return obj
}

func (s *Sheet) scanPicture(f *excelize.File, i int, j int) (reflect.Value, error) {
	var pictures []Picture
	var err error

	pics, err := f.GetPictures(s.sheet, cell(i+1, j+1))
	if err != nil {
		return reflect.Value{}, err
	}
	pictures = functools.Map(func(pic excelize.Picture) Picture {
		return Picture{
			File:     pic.File,
			Format:   (*PicFormat)(pic.Format),
			withPath: false,
		}
	}, pics)
	if len(pictures) != 0 {
		return reflect.ValueOf(pictures[0]), nil
	}
	return reflect.Value{}, nil
}

func (s *Sheet) validate(value any, valid string) error {
	if valid != "" {
		if value != nil {
			return validate.Var(value, valid)
		}
	}
	return nil
}

func (s *Sheet) marshaler(field any, value string) (bool, error) {
	if f, ok := field.(XLSXUnmarshaler); ok {
		if err := f.UnmarshalXLSX([]byte(value)); err != nil {
			return false, err
		}
		return true, nil
	} else if f, ok := field.(encoding.TextUnmarshaler); ok {
		if err := f.UnmarshalText([]byte(value)); err != nil {
			return false, err
		}
		return true, nil
	}
	return false, nil
}

func (s *Sheet) scanRow(f *excelize.File, t reflect.Type, row []string, obj map[string]string, o reflect.Value, i int, date1904 bool) error {
	for j := 0; j < t.NumField(); j++ {
		if !s.collectErrors && len(s.errors) > 0 {
			return s.errors[0]
		}
		tag := getFieldName(t.Field(j))
		valid := t.Field(j).Tag.Get("validate")
		field := o.Elem().Field(j).Addr().Interface()
		fieldInterface := o.Elem().Field(j).Interface()
		fieldType := reflect.TypeOf(field)
		value, ok := obj[tag]
		if !ok {
			if _, ok := fieldInterface.(Picture); !ok {
				continue
			}
			if rv, err := s.scanPicture(f, i, j); err != nil {
				s.errors = append(s.errors, Error{Row: Row{ID: i, Data: obj}, mesg: err.Error()})
			} else {
				o.Elem().Field(j).Set(rv)
			}
			continue
		}

		if needValidate, err := s.marshaler(field, value); err != nil {
			s.errors = append(s.errors, Error{Row: Row{ID: i, Data: obj}, mesg: fmt.Sprintf("%s: %s", tag, err.Error())})
			continue
		} else if needValidate {
			goto validate
		}

		if rv, err := getReflectValue(value, fieldType.Elem()); err == nil {
			if _, ok := fieldInterface.(time.Time); ok {
				// styleID := s.timeStyle(f, rv)
				col := 0
				functools.ForEach(row, func(elem string) {
					col++
					data, err := s.scanTime(f, col, i, elem, value, date1904)
					if err != nil {
						s.errors = append(s.errors, Error{Row: Row{ID: i, Data: obj}, mesg: fmt.Sprintf("%s: %s", tag, err.Error())})
					} else {
						o.Elem().Field(j).Set(reflect.ValueOf(data))
					}
				})
				goto validate
			}
			o.Elem().Field(j).Set(rv)
		} else {
			s.errors = append(s.errors, Error{Row: Row{ID: i, Data: obj}, mesg: fmt.Sprintf("%s: %s", tag, err.Error())})
			continue
		}
	validate:
		if err := s.validate(o.Elem().Field(j).Interface(), valid); err != nil {
			s.errors = append(s.errors, Error{Row: Row{ID: i, Data: obj}, mesg: fmt.Sprintf("%s: %s", tag, err.Error())})
		}
	}
	return nil
}

func (s *Sheet) scanSheet(f *excelize.File, rv reflect.Value) error {
	props, err := f.GetWorkbookProps()
	if err != nil {
		return err
	}
	date1904 := isDate1904(props)

	t := rv.Type().Elem().Elem()

	rows, err := f.GetRows(s.sheet)
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
	var data dataArray = newDataArray(length)
	n := 0
	for i, row := range rows {
		if i == 0 {
			schema = append(schema, functools.Map(func(s string) string { return strings.TrimSpace(s) }, row)...)
			continue
		}
		obj := createObjectMap(schema, row)
		if len(obj) == 0 {
			continue
		}
		data.indexArr = append(data.indexArr, i-1)
		n++
		o := reflect.New(t)
		if err := s.scanRow(f, t, row, obj, o, i, date1904); err != nil {
			return err
		}
		data.array[i-1] = o.Elem()
	}
	items := reflect.MakeSlice(rv.Type().Elem(), n, n)
	for i, index := range data.indexArr {
		items.Index(i).Set(data.array[index])
	}
	rv.Elem().Set(items)
	if len(s.errors) > 0 {
		return s.errors[0]
	}
	return nil
}

func (s *Sheet) Scan(v any) error {
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
	s.title = title
	s.colCnt = len(title)
	if s.useTextStyle {
		f.SetColStyle(sheet, "A:"+toTwentySix(s.colCnt), s.style)
	}
	for _, v := range title {
		f.SetCellStr(sheet, col(), v)
	}
}

func (s *Sheet) exportPic(f *excelize.File, pic Picture, col column) error {
	if pic.withPath {
		if err := f.AddPicture(s.sheet, col(), pic.Name, (*excelize.GraphicOptions)(pic.Format)); err != nil {
			return err
		}
	} else {
		if err := f.AddPictureFromBytes(s.sheet, col(),
			&excelize.Picture{
				File:   pic.File,
				Format: (*excelize.GraphicOptions)(pic.Format),
			}); err != nil {
			return err
		}
	}
	return nil
}

func (s *Sheet) exportCell(f *excelize.File, c Cell, col column) error {
	if c.HyperLink.Link != "" {
		column := col()
		f.SetCellStr(s.sheet, column, c.Value)
		f.SetCellHyperLink(s.sheet, column, c.HyperLink.Link, string(c.HyperLink.Type))
		if c.Style != nil {
			style, err := f.NewStyle(c.Style)
			if err != nil {
				return err
			}
			if err := f.SetCellStyle(s.sheet, column, column, style); err != nil {
				return err
			}
		}
	}
	return nil
}

func (s *Sheet) exportStruct(f *excelize.File, field any, col column) error {
	if pic, ok := field.(Picture); ok {
		return s.exportPic(f, pic, col)
	} else if cell, ok := field.(Cell); ok {
		return s.exportCell(f, cell, col)
	}
	if fun, ok := field.(XLSXMarshaler); ok {
		data, err := fun.MarshalXLSX()
		if err != nil {
			return err
		}
		f.SetCellStr(s.sheet, col(), string(data))
	} else if fun, ok := field.(encoding.TextMarshaler); ok {
		data, err := fun.MarshalText()
		if err != nil {
			return err
		}
		f.SetCellStr(s.sheet, col(), string(data))
	} else if t, ok := field.(time.Time); ok {
		f.SetCellValue(s.sheet, col(), t)
	} else {
		panic("struct type must implement MarshalXLSX or MarshalText")
	}
	return nil
}

func (s *Sheet) exportRow(f *excelize.File, obj reflect.Value, col column) error {
	t := obj.Type()
	for i := 0; i < obj.NumField(); i++ {
		field := obj.Field(i)
		fieldInterface := field.Interface()
		if field.Kind() == reflect.Struct {
			if err := s.exportStruct(f, fieldInterface, col); err != nil {
				return err
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
					colIdx := col()
					f.SetCellStr(s.sheet, colIdx, string(data))
					continue
				} else if fun, ok := fieldInterface.(encoding.TextMarshaler); ok {
					data, err := fun.MarshalText()
					if err != nil {
						return err
					}
					colIdx := col()
					f.SetCellStr(s.sheet, colIdx, string(data))
					continue
				}
				colIdx := col()
				f.SetCellStr(s.sheet, colIdx, toString(fieldInterface))
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

	sheet, err := f.NewSheet(s.sheet)
	if err != nil {
		return err
	}
	f.SetActiveSheet(sheet)

	s.exportTitle(f, s.filter, s.sheet, t, cellGenerator(1))

	slice := rv.Elem()

	if err := s.exportRows(f, slice); err != nil {
		return err
	}

	return nil
}

func (s *Sheet) export(f *excelize.File, v any) error {
	if s.sheet == "" {
		s.sheet = defaultSheet
	}
	rv := reflect.ValueOf(v)
	if rv.Kind() != reflect.Pointer || rv.IsNil() || rv.Type().Elem().Kind() != reflect.Slice {
		panic("param must be slice ptr")
	}
	if err := s.sheetExport(f, rv); err != nil {
		return err
	}
	if s.sheet != defaultSheet {
		f.DeleteSheet(defaultSheet)
	}
	return nil
}

// Export exports the sheet to a bytes.Buffer.
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

// ExportTo exports the sheet to a io.Writer.
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

// Filter sets the filter of the sheet.
func (s *Sheet) Filter(schema Schema) *Sheet {
	s.filter = schema
	return s
}
