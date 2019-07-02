package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"regexp"
	"strconv"

	"github.com/xuri/excelize"
)

func main() {
	var s []string
	var input, output string
	flag.StringVar(&input, "input", "123", "input")
	flag.StringVar(&output, "output", "123", "output")
	flag.Parse()
	s, _ = GetAllFile(input, s)
	// s, _ = GetAllFile(dir, s)

	xlsx := excelize.NewFile()
	sheetname := xlsx.NewSheet("Sheet1")
	err := xlsx.SetColWidth("Sheet1", "B", "C", 80)
	xlsx.SetActiveSheet(sheetname)
	xlsx.SetCellValue("Sheet1", "B1", "文件所在路径")
	xlsx.SetCellValue("Sheet1", "C1", "中文")
	xlsx.SetCellValue("Sheet1", "D1", "English")
	stylebold, err := xlsx.NewStyle(`{"alignment":{"horizontal":"center","Vertical":"center"},"font":{"bold":true,"family":"微软雅黑","size":12}}`)

	if err != nil {
		fmt.Println(err)
	}
	xlsx.SetCellStyle("Sheet1", "B1", "B1", stylebold)
	xlsx.SetCellStyle("Sheet1", "C1", "C1", stylebold)
	xlsx.SetCellStyle("Sheet1", "D1", "D1", stylebold)

	i := 2
	for _, filename := range s {
		tail := path.Ext(path.Base(filename))

		if tail == ".html" {

			htmlch, _ := regexCh(filename)

			for _, values := range htmlch {
				xlsx.SetCellValue("Sheet1", "B"+strconv.Itoa(i), extra(filename, input))
				xlsx.SetCellValue("Sheet1", "C"+strconv.Itoa(i), values)

				i = i + 1

			}

		} else if tail == ".htm" {
			// fmt.Println(filename)
			htmChi, _ := regexCh(filename)
			for _, values := range htmChi {
				xlsx.SetCellValue("Sheet1", "B"+strconv.Itoa(i), extra(filename, input))
				xlsx.SetCellValue("Sheet1", "C"+strconv.Itoa(i), values)

				i = i + 1

			}
		} else if tail == ".js" {
			// fmt.Println(filename)
			jsChin, _ := regexCh(filename)
			for _, values := range jsChin {
				xlsx.SetCellValue("Sheet1", "B"+strconv.Itoa(i), extra(filename, input))
				xlsx.SetCellValue("Sheet1", "C"+strconv.Itoa(i), values)
				// fmt.Println("C"+strconv.Itoa(i), values)
				// _, _ = index, values
				i = i + 1
				// fmt.Println("htm i:", i)
			}
		} else {

		}

	}

	errs := xlsx.SaveAs(output + ".xlsx")
	if errs != nil {
		fmt.Println(errs)
	}
}

//RemoveRepeatedElement 数组去重
func RemoveRepeatedElement(arr []string) (newArr []string) {
	newArr = make([]string, 0)
	for i := 0; i < len(arr); i++ {
		repeat := false
		for j := i + 1; j < len(arr); j++ {
			if arr[i] == arr[j] {
				repeat = true
				break
			}
		}
		if !repeat {
			newArr = append(newArr, arr[i])
		}
	}
	return
}

/*
//getFilelist 遍历文件夹和文件
func getFilelist(path string) {
	err := filepath.Walk(path, func(path string, f os.FileInfo, err error) error {
		if f == nil {
			return err
		}
		if f.IsDir() {
			return nil
		}
		println(path)
		return nil
	})
	if err != nil {
		fmt.Printf("filepath.Walk() returned %v\n", err)
	}
}
*/
// GetAllFile 遍历文件
func GetAllFile(pathname string, s []string) ([]string, error) {
	rd, err := ioutil.ReadDir(pathname)
	if err != nil {
		fmt.Println("read dir fail:", err)
		return s, err
	}
	for _, fi := range rd {
		if fi.IsDir() {
			fullDir := pathname + "/" + fi.Name()
			s, err = GetAllFile(fullDir, s)
			if err != nil {
				fmt.Println("read dir fail:", err)
				return s, err
			}
		} else {
			fullName := pathname + "/" + fi.Name()
			s = append(s, fullName)
		}
	}
	return s, nil
}
func regexCh(str string) ([]string, int) {
	loginfile, err := os.Open(str)
	defer loginfile.Close()
	if err != nil {
		fmt.Println(err)
	}
	buf := make([]byte, 102400)
	for {
		//Read函数会改变文件当前偏移量
		len, _ := loginfile.Read(buf)

		//读取字节数为0时跳出循环
		if len == 0 {
			break
		}
		//	fmt.Println(string(buf))
	}
	res := string(buf)
	//`[\p{Han}]+`     ^[\u4e00-\u9fa5]$ reflect.ValueOf
	reg := regexp.MustCompile(`[\p{Han}]+`)
	data := reg.FindAllString(res, -1)
	newdata := RemoveRepeatedElement(data)
	return newdata, len(newdata)
}

func extra(str, str1 string) string {
	reg := regexp.MustCompile(str1)
	s := reg.ReplaceAllString(str, "")
	return s
}
