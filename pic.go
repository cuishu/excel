package excel

import (
	"bytes"
	"encoding/json"
	"errors"
	"reflect"

	"github.com/gabriel-vasile/mimetype"
)

type PicFormat struct {
	XScale float64 `json:"x_scale"`
	YScale float64 `json:"y_scale"`
}

func (format PicFormat) String() string {
	if format.XScale == 0 && format.YScale == 0 {
		return ""
	}
	js, _ := json.Marshal(&format)
	return string(js)
}

type Picture struct {
	Name     string
	File     []byte
	Format   PicFormat
	withPath bool
}

func getPicExtName(mime string) (string, error) {
	switch mime {
	case "image/jpeg":
		return ".jpeg", nil
	case "image/png":
		return ".png", nil
	default:
		return "", errors.New("invalid image type: must be jpeg or png")
	}
}

func NewPicture(path string, format PicFormat) Picture {
	return Picture{
		Name:     path,
		Format:   format,
		withPath: true,
	}
}

func NewPictureFromBytes(file []byte, format PicFormat) (Picture, error) {
	extName, err := getPicExtName(mimetype.Detect(file).String())
	if err != nil {
		return Picture{}, err
	}
	return Picture{
		Name:     extName,
		File:     file,
		Format:   format,
		withPath: false,
	}, nil
}

func (pic Picture) Buffer() *bytes.Buffer {
	return bytes.NewBuffer(pic.File)
}

var picReflectType = reflect.TypeOf(Picture{})
