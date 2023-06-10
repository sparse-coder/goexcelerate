package goexcelerate

import (
	"archive/zip"
	"embed"
	"fmt"
	"text/template"
	"os"
	"time"
)

//go:embed all:templates
var templateAssets embed.FS

type Writer struct {
	DocPropsAppTemplate  *template.Template
	DocPropsCoreTemplate *template.Template
	ContentTypesTemplate *template.Template
	RelTemplate          *template.Template

	WorkbookTemplate     *template.Template
	WorkbookRelsTemplate *template.Template
	WorksheetTemplate    *template.Template
	VbaProjectBinFile    string
	PWorkbook            *Workbook
}

func NewWriter() *Writer {
	writer := &Writer{
		DocPropsAppTemplate:  template.Must(template.ParseFS(templateAssets, "templates/docProps/app.xml")),
		DocPropsCoreTemplate: template.Must(template.ParseFS(templateAssets, "templates/docProps/core.xml")),
		ContentTypesTemplate: template.Must(template.ParseFS(templateAssets, "templates/Content_Types.xml")),
		RelTemplate:          template.Must(template.ParseFS(templateAssets, "templates/_rels/.rels")),
		WorkbookTemplate:     template.Must(template.ParseFS(templateAssets, "templates/xl/workbook.xml")),
		WorkbookRelsTemplate: template.Must(template.ParseFS(templateAssets, "templates/xl/_rels/workbook.xml.rels")),
		WorksheetTemplate:    template.Must(template.ParseFS(templateAssets, "templates/xl/worksheets/sheet.xml")),
		//VbaProjectBinFile:    "",
	}
	return writer
}

func (wr *Writer) Save(fileHandle string) {
	archive, err := os.Create(fileHandle)
	if err != nil {
		panic(err)
	}
	defer archive.Close()
	zipWriter := zip.NewWriter(archive)

	docWr, err := zipWriter.Create("docProps/app.xml")
	if err != nil {
		panic(err)
	}
	wr.DocPropsAppTemplate.Execute(docWr, wr.PWorkbook)

	coreWr, err := zipWriter.Create("docProps/core.xml")
	if err != nil {
		panic(err)
	}
	crDate := struct {
		Date string
	}{Date: time.Now().Format(time.DateTime)}
	wr.DocPropsCoreTemplate.Execute(coreWr, crDate)

	contentTypesWr, err := zipWriter.Create("[Content_Types].xml")
	if err != nil {
		panic(err)
	}
	wr.ContentTypesTemplate.Execute(contentTypesWr, wr.PWorkbook)

	relsWr, err := zipWriter.Create("_rels/.rels")
	if err != nil {
		panic(err)
	}
	wr.RelTemplate.Execute(relsWr, true)

	wbWr, err := zipWriter.Create("xl/workbook.xml")
	if err != nil {
		panic(err)
	}
	wr.WorkbookTemplate.Execute(wbWr, wr.PWorkbook)

	wbRelsWr, err := zipWriter.Create("xl/_rels/workbook.xml.rels")
	if err != nil {
		panic(err)
	}
	wr.WorkbookRelsTemplate.Execute(wbRelsWr, wr.PWorkbook)

	// write all the worksheets
	for index, sheet := range wr.PWorkbook.Worksheets {
		fname := fmt.Sprintf("xl/worksheets/sheet%d.xml", index)
		tWr, err := zipWriter.Create(fname)

		if err != nil {
			panic(err)
		}
		wr.WorksheetTemplate.Execute(tWr, sheet)
	}

	zipWriter.Close()
}
