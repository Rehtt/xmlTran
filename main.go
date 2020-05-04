package main

import (
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io"
	"os"
	"regexp"
	"strconv"
)

func main() {
	//xml2xlsx()
	//xlsx2xml()
}

func xlsx2xml() {
	f, err := excelize.OpenFile("strings.xlsx")
	if err != nil {
		fmt.Println("请确保文件名为strings.xlsx")
		return
	}
	out, err := os.Create("strings.xml")
	if err != nil || err == io.EOF {
		fmt.Println("保存文件失败，请检查是否有写入权限")
	}
	defer out.Close()
	out.WriteString("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n")
	for i, v := range f.GetRows("Sheet1") {
		if i != 0 {
			name := v[0]
			value := v[2]
			if value == "" {
				value = v[1]

			}
			if value != "" {
				out.WriteString(fmt.Sprintf("<string name=\"%s\">%s</string>\n", name, value))
			} else {
				out.WriteString(name + "\n")
			}

		}
	}
	out.WriteString("</resources>")

}
func xml2xlsx() {
	file, _ := os.OpenFile("strings.xml", os.O_RDONLY, 0)
	defer file.Close()
	res := bufio.NewScanner(file)
	f := excelize.NewFile()
	index := f.NewSheet("Sheet1")
	f.SetCellValue("Sheet1", "A1", "NAME")
	f.SetCellValue("Sheet1", "B1", "原文")
	f.SetCellValue("Sheet1", "C1", "翻译")

	i := 1

	name := regexp.MustCompile(`<string name="(.*?)">`)
	value := regexp.MustCompile(`">(.*?)</string>`)
	note := regexp.MustCompile(`<!--.*`)
	notend := regexp.MustCompile(`.*-->`)
	notea := regexp.MustCompile(`<!--.*-->`)
	for res.Scan() {
		namev := name.FindStringSubmatch(res.Text())
		valuev := value.FindStringSubmatch(res.Text())
		if len(namev) != 0 && len(valuev) != 0 {
			i++
			f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), namev[1])
			f.SetCellValue("Sheet1", "B"+strconv.Itoa(i), valuev[1])
		} else if noteav := notea.FindStringSubmatch(res.Text()); len(noteav) != 0 {
			i++
			f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), noteav[0])
		} else if notev := note.FindStringSubmatch(res.Text()); len(notev) != 0 {
			i++
			f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), notev[0]+"\n")
		} else if notendv := notend.FindStringSubmatch(res.Text()); len(notendv) != 0 {
			f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), f.GetCellValue("Sheet1", "A"+strconv.Itoa(i))+notendv[0])
		}
	}
	f.SetActiveSheet(index)
	if err := f.SaveAs("strings.xlsx"); err != nil || err == io.EOF {
		fmt.Println("保存文件失败，请检查是否有写入权限")
		return
	}
}
