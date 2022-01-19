package excel

import (
	"fmt"
	"io/ioutil"
	"os"
	"testing"
)

type Human struct {
	ID   int    `xlsx:"id"`
	Name string `xlsx:"name"`
}

type Animal struct {
	ID   int    `xlsx:"id"`
	Name string `xlsx:"name"`
}

type ExcelExample struct {
	Human  []Human  `xlsx:"Human"`
	Animal []Animal `xlsx:"Animal"`
}

func TestExport(t *testing.T) {
	example := ExcelExample{
		Human:  []Human{{1, "Smith"}},
		Animal: []Animal{{1, "Wolverine"}},
	}
	_, err := Excel{}.Export(&example)
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
}

func TestScan(t *testing.T) {
	buff, err := Excel{}.Export(&ExcelExample{
		Human:  []Human{{1, "Smith"}},
		Animal: []Animal{{1, "Wolverine"}},
	})
	if err != nil {
		t.Log(err)
		t.FailNow()
	}
	ioutil.WriteFile("b.xlsx", buff.Bytes(), 0644)

	var example ExcelExample
	if err := (Excel{Filename: "b.xlsx"}).Scan(&example); err != nil {
		t.Error(err)
		t.FailNow()
	}
	if len(example.Animal) == 0 || len(example.Human) == 0 {
		t.Fail()
	}
	fmt.Println(example.Human)
	fmt.Println(example.Animal)
	os.Remove("b.xlsx")
}
