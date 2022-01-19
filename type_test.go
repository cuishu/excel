package excel

import (
	"fmt"
	"testing"
)

func TestTwentySix(t *testing.T) {
	for i := 1; i < 999; i++ {
		fmt.Println(toTwentySix(i))
	}
	t.Fail()
}
