package goexcelerate

var COORD_TO_COLUMN map[int]string

func init() {
	COORD_TO_COLUMN = GenerateColumnMappings()
}

func GenerateColumnMappings() map[int]string {
	mappings := make(map[int]string)
	for i := 1; i <= 18278; i++ {
		column := ""
		for n := i; n > 0; n = (n - 1) / 26 {
			column = string('A'+(n-1)%26) + column
		}
		mappings[i] = column
	}
	return mappings
}