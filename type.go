package excel

import (
	"fmt"
	"reflect"
	"strconv"

	"github.com/go-playground/validator/v10"
)

type Int interface {
	int | int8 | int16 | int32 | int64 |
	uint | uint8 | uint16 | uint32 | uint64
}

func parseInt[T Int](s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseInt(s, 10, 8)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(T(v)), nil
}

type Float interface {
	float32 | float64
}

func parseFloat[T Float](s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(T(v)), nil
}

func parseBool(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseBool(s)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(v), nil
}

func getReflectValue(s string, t reflect.Type) (reflect.Value, error) {
	var rv reflect.Value
	switch t.Kind() {
	case reflect.String:
		return reflect.ValueOf(s), nil
	case reflect.Int8:
		return parseInt[int8](s)
	case reflect.Int16:
		return parseInt[int16](s)
	case reflect.Int32:
		return parseInt[int32](s)
	case reflect.Int64:
		return parseInt[int64](s)
	case reflect.Int:
		return parseInt[int](s)
	case reflect.Uint8:
		return parseInt[uint8](s)
	case reflect.Uint16:
		return parseInt[uint16](s)
	case reflect.Uint32:
		return parseInt[uint32](s)
	case reflect.Uint64:
		return parseInt[uint64](s)
	case reflect.Uint:
		return parseInt[uint](s)
	case reflect.Float32:
		return parseFloat[float32](s)
	case reflect.Float64:
		return parseFloat[float64](s)
	case reflect.Bool:
		return parseBool(s)
	}
	return rv, nil
}

func toString(v any) string {
	if bv, ok := v.([]byte); ok {
		return string(bv)
	}
	return fmt.Sprintf("%v", v)
}

func isTime(rt reflect.Type) bool {
	if rt.PkgPath() == "time" && rt.Name() == "Time" {
		return true
	}
	return false
}

var validate = validator.New()

var twentySixTable = []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func toTwentySix(n int) string {
	var str string
	var k int
	var temp []int = make([]int, 0, 3)
	if n > 26 {
		for {
			k = n % 26
			if k == 0 {
				temp = append(temp, 26)
				k = 26
			} else {
				temp = append(temp, k)
			}
			n = (n - k) / 26
			if n <= 26 {
				temp = append(temp, n)
				break
			}
		}
	} else {
		return twentySixTable[n]
	}
	for _, v := range temp {
		str = twentySixTable[v] + str
	}
	return str
}
