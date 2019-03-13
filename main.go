package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var excel = make([]string, 2)

func main() {

	cells, err := readConfig("cfg.conf")
	folder := "C:\\_test"

	err = iterateFolder2(folder)
	if err != nil {
		log.Fatal(err)
	}

	for _, ff := range excel {
		if ff != "" {
			log.Println(ff)

			_, err := readExcel(ff, cells)
			if err != nil {
				log.Fatal(err)
			}
		}
	}

}

func iterateFolder(folder string) error {

	contentFolder, err := ioutil.ReadDir(folder)
	if err != nil {
		return err
	}

	for _, f := range contentFolder {
		fmt.Println(f.Name())
	}

	return nil
}

func iterateFolder2(folder string) error {

	os.Chdir(folder)

	subDirToSkip := "skip"

	err := filepath.Walk(".", func(path string, info os.FileInfo, err error) error {
		if err != nil {
			log.Println(err)
			return err
		}
		if info.IsDir() && info.Name() == subDirToSkip {
			fmt.Printf("skipped: %+v \n", info.Name())
			return filepath.SkipDir
		}
		if !info.IsDir() {
			if strings.HasSuffix(path, ".xlsx") || strings.HasSuffix(path, ".xls") {
				fmt.Printf("excel: %q\n", folder+"\\"+path)
				excel = append(excel, folder+"\\"+path)
			} else {
				fmt.Printf("file: %q\n", folder+"\\"+path)
			}
		}

		return nil
	})
	if err != nil {
		log.Println("error walking the path %q: %v\n", folder, err)
		return err
	}

	return nil
}

func readExcel(file string, cells []string) ([]string, error) {

	data := make([]string, len(cells)+2)
	data[0] = path.Dir(file)
	data[1] = path.Base(file)

	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		return nil, err
	}

	for i, c := range cells {
		data[i+2] = xlsx.GetCellValue("Sheet1", c)
	}

	fmt.Println(data)
	return data, nil
}

func readConfig(file string) ([]string, error) {
	content, err := ioutil.ReadFile(file)
	if err != nil {
		return nil, err
	}

	return strings.Split(string(content), ";"), nil

}
