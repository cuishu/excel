package excel

import (
	"fmt"
	"reflect"
	"strconv"

	"github.com/go-playground/validator/v10"
)

func parseInt8(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseInt(s, 10, 8)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(int8(v)), nil
}

func parseInt16(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseInt(s, 10, 16)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(int16(v)), nil
}

func parseInt32(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseInt(s, 10, 32)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(int32(v)), nil
}

func parseInt64(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseInt(s, 10, 64)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(int64(v)), nil
}

func parseInt(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseInt(s, 10, 64)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(int(v)), nil
}

func parseUint8(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseUint(s, 10, 8)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(uint8(v)), nil
}

func parseUint16(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseUint(s, 10, 16)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(uint16(v)), nil
}

func parseUint32(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseUint(s, 10, 32)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(uint32(v)), nil
}

func parseUint64(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseUint(s, 10, 64)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(uint64(v)), nil
}

func parseUint(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseUint(s, 10, 64)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(uint(v)), nil
}

func parseFloat32(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(float32(v)), nil
}

func parseFloat64(s string) (reflect.Value, error) {
	var rv reflect.Value
	v, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return rv, err
	}
	return reflect.ValueOf(float64(v)), nil
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
		return parseInt8(s)
	case reflect.Int16:
		return parseInt16(s)
	case reflect.Int32:
		return parseInt32(s)
	case reflect.Int64:
		return parseInt64(s)
	case reflect.Int:
		return parseInt(s)
	case reflect.Uint8:
		return parseUint8(s)
	case reflect.Uint16:
		return parseUint16(s)
	case reflect.Uint32:
		return parseUint32(s)
	case reflect.Uint64:
		return parseUint64(s)
	case reflect.Uint:
		return parseUint(s)
	case reflect.Float32:
		return parseFloat32(s)
	case reflect.Float64:
		return parseFloat64(s)
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
