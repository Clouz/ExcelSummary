package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var excel = make([]string, 2)
var data = make([][]string, 2)

func main() {

	cells, sheet, err := readConfig("cfg.conf")
	if err != nil {
		log.Fatalln(err)
		fmt.Scanln() // wait for Enter Key
	}

	if len(os.Args) != 2 {
		log.Fatalln(err)
		fmt.Scanln() // wait for Enter Key
	}
	folder := os.Args[1]
	//folder := "C:\\Users\\sacoa002036\\Desktop\\CS_Aramco"

	log.Println("Scanning Sub Directory... ", folder)
	err = iterateFolder(folder)
	if err != nil {
		log.Fatal(err)
		fmt.Scanln() // wait for Enter Key
	}
	log.Println("Scanning Done")

	fmt.Println("Reading excel content...")
	for _, ff := range excel {
		if ff != "" {
			log.Println(ff)

			d, err := readExcel(ff, cells, sheet)
			data = append(data, d)

			if err != nil {
				log.Fatal(err)
				fmt.Scanln() // wait for Enter Key
			}
		}
	}
	log.Println("Reading Done")

	err = writeExcel(data, folder)
	if err != nil {
		log.Fatalln(err)
		fmt.Scanln() // wait for Enter Key
	}

	fmt.Scanln() // wait for Enter Key
}

func iterateFolder_hold(folder string) error {

	contentFolder, err := ioutil.ReadDir(folder)
	if err != nil {
		return err
	}

	for _, f := range contentFolder {
		fmt.Println(f.Name())
	}

	return nil
}

func iterateFolder(folder string) error {

	dirCount := 0
	fileCount := 0
	excelCount := 0

	os.Chdir(folder)

	err := filepath.Walk(".", func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			dirCount++
		}

		if !info.IsDir() {
			p := strings.ToLower(path)
			if strings.HasSuffix(p, ".xlsx") || strings.HasSuffix(p, ".xls") {
				excel = append(excel, folder+"\\"+path)
				excelCount++
			} else {
				fileCount++
			}
		}

		return nil
	})
	if err != nil {
		return err
	}

	fmt.Println("Scanned directory: ", dirCount, "\nScanned Excel: ", excelCount, "\nScanned File: ", fileCount)
	return nil
}

func readExcel(file string, cells []string, sheet string) ([]string, error) {

	data := make([]string, len(cells)+2)
	data[0], data[1] = filepath.Split(file)

	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		return nil, err
	}
	//maps := xlsx.GetSheetMap()

	for i, c := range cells {
		data[i+2] = xlsx.GetCellValue("3doCS", c)
	}

	//fmt.Println(data)
	return data, nil
}

func readConfig(file string) ([]string, string, error) {
	content, err := ioutil.ReadFile(file)
	if err != nil {
		return nil, "", err
	}

	sheet := strings.Split(string(content), ":")
	cells := strings.Split(sheet[1], ";")
	return cells, sheet[0], nil

}

func writeExcel(data [][]string, folder string) error {
	xlsx := excelize.NewFile()

	for irow, row := range data {
		for icol, collum := range row {
			cell := fmt.Sprint(excelize.ToAlphaString(icol), irow)
			xlsx.SetCellValue("Sheet1", cell, collum)
		}
	}
	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(1)
	// Save xlsx file by the given path.
	err := xlsx.SaveAs(folder + "\\Summary.xlsx")
	if err != nil {
		return err
	}
	return nil
}
