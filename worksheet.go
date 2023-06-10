package goexcelerate

import "fmt"

type Worksheet struct {
	Name          string
	Cells         map[int]map[int]any
	ParentWb      *Workbook
	ShowGridLines bool
	Panes         *Panes
}

func (ws *Worksheet) SetCellValue(x int, y int, value any) {
	if x <= 0 && y <= 0 {
		panic("x and y should be >=1")
	}
	if ws.Cells == nil {
		ws.Cells = make(map[int]map[int]any)
	}
	if ws.Cells[x] == nil {
		ws.Cells[x] = make(map[int]any)
	}
	ws.Cells[x][y] = value
}

func (ws Worksheet) GetXMLData() <-chan string {
	//rows := []string{}
	channel := make(chan string)
	go func(){
		defer close(channel)
		for row, cols := range ws.Cells {
			GetRowString(row, cols, channel)
		}
	}()
	return channel
}

func GetRowString(row int, cols map[int]any, ch chan<-string) {
	row_ := fmt.Sprintf(`<row r="%d"> `, row)
	for col, val := range cols {
		row_ = row_ + GetColString(row, col, val)
	}
	row_ = row_ + "</row>"
	ch <- row_
}

func GetColString(row int, col int, val any) string {
	column := fmt.Sprintf(`<c r="%s%d"> <v>%v</v> </c>`, COORD_TO_COLUMN[col], row, val)
	return column
}
