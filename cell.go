package excel

import (
	excelize "github.com/xuri/excelize/v2"
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
