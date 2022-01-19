package excel

import (
	"reflect"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

type LinkType string

const (
	External LinkType = "External"
	Location LinkType = "Location"
)

type HyperLink struct {
	Link string
	Type LinkType
}

type Cell struct {
	Value     string
	HyperLink HyperLink
	Style     *excelize.Style
}

var cellReflectType = reflect.TypeOf(Cell{})
