package excel

import (
	"fmt"
	"reflect"
	"testing"
	"time"
)

func TestTwentySix(t *testing.T) {
	for i := 1; i < 999; i++ {
		fmt.Println(toTwentySix(i))
	}
	t.Fail()
}

func TestType(t *testing.T) {
	now := time.Now()
	rt := reflect.TypeOf(now)
	fmt.Println(rt.Name())
	fmt.Println(rt.PkgPath())
	t.Fail()
}
