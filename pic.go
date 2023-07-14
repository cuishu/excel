package excel

import (
	"bytes"
	"errors"
	"reflect"

	"github.com/gabriel-vasile/mimetype"
	"github.com/xuri/excelize/v2"
)

type PicFormat excelize.GraphicOptions

type Picture struct {
	Name     string
	File     []byte
	Format   *PicFormat
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

func NewPicture(path string, format *PicFormat) Picture {
	return Picture{
		Name:     path,
		Format:   format,
		withPath: true,
	}
}

func NewPictureFromBytes(file []byte, format *PicFormat) (Picture, error) {
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
