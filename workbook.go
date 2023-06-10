package goexcelerate

type Workbook struct {
	Worksheets []*Worksheet
	Encoding   string
	HasMacros  bool
	HasStyles  bool
	Writer     *Writer
}

func NewWorkbook() *Workbook {
	wb := &Workbook{
		Encoding: "utf-8",
		Writer:   NewWriter(),
	}
	wb.Writer.PWorkbook = wb
	return wb
}

func (wb *Workbook) AddSheet(w *Worksheet) {
	for _, sheet := range wb.Worksheets {
		if sheet.Name == w.Name {
			panic("Duplicate worksheet names are not permitted.")
		}
	}
	w.ParentWb = wb
	wb.Worksheets = append(wb.Worksheets, w)
}

func (wb *Workbook) Len() int {
	return len(wb.Worksheets)
}

type XMLData struct {
	Idx int
	Ws  Worksheet
}

func (wb *Workbook) GetXMLData() []*Worksheet {
	return wb.Worksheets
}

func (wb *Workbook) Save(fileHandle string) {
	wb.Writer.Save(fileHandle)
}
