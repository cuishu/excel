package excel

import (
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"
	"testing"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

const (
	// Male 男
	Male = 1
	// Female 女
	Female = 2
)

type Sex int

func (sex Sex) MarshalXLSX() ([]byte, error) {
	switch sex {
	case Male:
		return []byte("男"), nil
	case Female:
		return []byte("女"), nil
	}
	return nil, errors.New("性别错误")
}

func (sex *Sex) UnmarshalXLSX(data []byte) error {
	s := string(data)
	s = strings.Trim(s, "\"")
	if s == "男" {
		*sex = Male
		return nil
	}
	if s == "女" {
		*sex = Female
		return nil
	}

	v, err := strconv.ParseUint(s, 10, 64)
	if err != nil {
		return err
	}
	if v == Male || v == Female {
		*sex = Sex(v)
		return nil
	}
	return errors.New("invalid value")
}

type Time struct {
	time.Time
}

func (t Time) MarshalXLSX() ([]byte, error) {
	return []byte(t.Format("2006-01-02 15:04:05")), nil
}

func (t *Time) UnmarshalXLSX(data []byte) error {
	var err error
	t.Time, err = time.ParseInLocation("2006-01-02 15:04:05", string(data), time.Local)
	return err
}

type TestObject struct {
	Name      string  `xlsx:"name" binding:"required"`
	Sex       Sex     `xlsx:"sex" binding:"oneof=1 2"`
	Age       int     `xlsx:"age"`
	Time      Time    `xlsx:"time"`
	Pic       Picture `xlsx:"picture"`
	HyperLink Cell    `xlsx:"hyperLink"`
}

func TestExportAndScan(t *testing.T) {
	e := &Sheet{filename: "a.xlsx", sheet: "Sheet3"}
	e.UseTextStyle()
	buff, err := e.Export(&[]TestObject{
		{"Smith", Male, 10, Time{time.Now()}, NewPicture("a.png", nil), Cell{}},
	})
	if err != nil {
		fmt.Println(err.Error())
		t.FailNow()
	}
	os.WriteFile("a.xlsx", buff.Bytes(), 0644)
	var ss []TestObject
	e.Scan(&ss)
	for _, s := range ss {
		fmt.Println(s.Name, s.Sex)
	}
	// os.Remove("a.xlsx")
	t.Fail()
}

func TestPicWithURL(t *testing.T) {
	url := "a.png"
	pic := NewPicture(url, nil)
	var objs []TestObject = make([]TestObject, 2)
	objs[0] = TestObject{Name: "嬴政", Sex: Male, Age: 25, Time: Time{Time: time.Now()}, Pic: pic, HyperLink: Cell{
		Value: "123",
		HyperLink: HyperLink{
			Link: "excel.go",
			Type: Location,
		},
		Style: &excelize.Style{
			Font: &excelize.Font{
				Color:     "#1265BE",
				Underline: "single",
			},
		},
	}}
	objs[1] = TestObject{Name: "刘彻", Sex: Male, Age: 24, Time: Time{Time: time.Now()}, Pic: pic, HyperLink: Cell{
		Value: "123",
		HyperLink: HyperLink{
			Link: "excel.go",
			Type: Location,
		}}}
	buff, err := (&Sheet{sheet: "Sheet1"}).Export(&objs)
	if err != nil {
		t.Error(err)
		t.FailNow()
	}
	os.WriteFile("a.xlsx", buff.Bytes(), 0644)
	os.Remove("a.xlsx")
}

func TestSheetScanFromReader(t *testing.T) {
	e := &Sheet{sheet: "Sheet3"}
	buff, err := e.Export(&[]TestObject{
		{"Smith", Male, 10, Time{time.Now()}, NewPicture("a.png", nil), Cell{}},
	})
	if err != nil {
		fmt.Println(err.Error())
		t.FailNow()
	}
	os.WriteFile("b.xlsx", buff.Bytes(), 0644)
	defer os.Remove("b.xlsx")

	file, err := os.Open("b.xlsx")
	if err != nil {
		t.FailNow()
	}
	defer file.Close()
	var data []TestObject
	xlsx := NewSheetFromReader(file, e.sheet)
	if err = xlsx.Scan(&data); err != nil {
		t.FailNow()
	}
	fmt.Println(data)

	xlsx = NewSheetFromReader(nil, e.sheet)
	if err = xlsx.Scan(&data); err == nil {
		t.FailNow()
	}
}

func TestOffset(t *testing.T) {
	e := Sheet{filename: "a.xlsx", sheet: "Sheet3"}
	var ss []TestObject
	e.Offset(1).Scan(&ss)
	for _, s := range ss {
		fmt.Println(s.Name, s.Sex)
	}
	t.Fail()
}

type TestTimeObject struct {
	Name string    `xlsx:"name" binding:"required"`
	Sex  Sex       `xlsx:"sex" binding:"oneof=1 2"`
	Age  int       `xlsx:"age"`
	Time time.Time `xlsx:"time"`
}

func TestTime(t *testing.T) {
	e := Sheet{filename: "a.xlsx", sheet: "Sheet3"}
	var ss []TestTimeObject
	e.Scan(&ss)
	for _, s := range ss {
		fmt.Println(s.Name, s.Sex, s.Time)
	}
	data, _ := e.Export(&ss)
	os.WriteFile("b.xlsx", data.Bytes(), 0644)
	t.Fail()
}
