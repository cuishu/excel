package excel

import (
	"fmt"
	"reflect"
	"strconv"

	"github.com/go-playground/validator/v10"
)

func getReflectValue(s string, t reflect.Type) (reflect.Value, error) {
	var rv reflect.Value
	switch t.Kind() {
	case reflect.String:
		rv = reflect.ValueOf(s)
	case reflect.Int8:
		v, err := strconv.ParseInt(s, 10, 8)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(int8(v))
	case reflect.Int16:
		v, err := strconv.ParseInt(s, 10, 16)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(int16(v))
	case reflect.Int32:
		v, err := strconv.ParseInt(s, 10, 32)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(int32(v))
	case reflect.Int64:
		v, err := strconv.ParseInt(s, 10, 64)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(int64(v))
	case reflect.Int:
		v, err := strconv.ParseInt(s, 10, 64)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(int(v))

	case reflect.Uint8:
		v, err := strconv.ParseUint(s, 10, 8)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(uint8(v))
	case reflect.Uint16:
		v, err := strconv.ParseUint(s, 10, 16)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(uint16(v))
	case reflect.Uint32:
		v, err := strconv.ParseUint(s, 10, 32)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(uint32(v))
	case reflect.Uint64:
		v, err := strconv.ParseUint(s, 10, 64)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(uint64(v))
	case reflect.Uint:
		v, err := strconv.ParseUint(s, 10, 64)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(uint(v))

	case reflect.Float32:
		v, err := strconv.ParseFloat(s, 32)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(float32(v))
	case reflect.Float64:
		v, err := strconv.ParseFloat(s, 64)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(float64(v))

	case reflect.Bool:
		v, err := strconv.ParseBool(s)
		if err != nil {
			return rv, err
		}
		rv = reflect.ValueOf(v)
	}
	return rv, nil
}

func toString(v interface{}) string {
	if bv, ok := v.([]byte); ok {
		return string(bv)
	}
	return fmt.Sprintf("%v", v)
}

var validate = validator.New()

var twentySixTable = []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func toTwentySix(n int) string {
	var str string
	var k int
	var temp []int
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
