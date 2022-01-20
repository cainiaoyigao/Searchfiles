package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var pathStr = ""
var queryType = 0 // 0 全民查询 1 模糊查询

func getPath(strPatch string) string {
	if strPatch == "" {
		strPatch = GetCurrentDirectory()
	}
	return strPatch
}

func GetCurrentDirectory() string {
	//返回绝对路径 filepath.Dir(os.Args[0])去除最后一个元素的路径
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		log.Fatal(err)
	}

	//将\替换成/
	return strings.Replace(dir, "\\", "/", -1)
}

func GetAllFile(pathname string, s []string) ([]string, error) {
	rd, err := ioutil.ReadDir(pathname)
	if err != nil {
		fmt.Println("read dir fail:", err)
		return s, err
	}

	for _, fi := range rd {
		if !fi.IsDir() {
			fullName := pathname + "/" + fi.Name()
			s = append(s, fullName)
		}
	}
	return s, nil
}

func main() {
	fmt.Println("........目前不支持指定地址查找，请将该程序放入查找的文件夹...........")
	fmt.Println("........目前只会查询当前文件夹下的xlsx文件，不会查询子文件夹下文件...........")
	pathdDirectory := getPath(pathStr)
	fmt.Println("........当前文件夹", pathdDirectory, "...........")
LOOP:
	if !selectType() {
		return
	}

	for {
		targetStr := FuzzyHanding()

		if targetStr == "exit" {
			goto LOOP
		}
		var files []string
		files, _ = GetAllFile(pathdDirectory, files)
		for _, filePath := range files {
			log.Println("filePath", filePath)
			fileType := path.Ext(filePath)
			if fileType != ".xlsx" {
				continue
			}
			ImportFromXLS(filePath, targetStr)
		}
	}
}

func selectType() bool {
	fmt.Println("\n请选择查找模式: 0全称 1模糊")
	var qType string
	_, err := fmt.Scan(&qType)
	qTypeInt, err := strconv.Atoi(qType)
	if err != nil {
		log.Println("选择错误")
		return false
	}
	queryType = qTypeInt

	return true
}

func FuzzyHanding() string {
	fmt.Println("\n请输入查找内容 输入exit重新选择查询方式:")
	var targetStr string
	_, err := fmt.Scan(&targetStr)
	if err != nil {
		return err.Error()
	}
	return targetStr
}

func ImportFromXLS(file, targetStr string) {
	f, err := excelize.OpenFile(file)
	if err != nil {
		log.Println(err)
		return
	}

	sheets := f.GetSheetMap()
	for _, sheet := range sheets {
		//log.Println("table name", sheet)
		rows := f.GetRows(sheet)
		for _, row := range rows {
			for _, colCell := range row {
				if IsQuery(targetStr, colCell) {
					log.Println("....................find path", file, "table ", sheet, "..............................")
				}
			}
		}
	}
}

func IsQuery(targetStr, colCell string) bool {
	if queryType == 0 {
		if targetStr == colCell {
			return true
		}
	} else if queryType == 1 {
		if strings.Contains(colCell, targetStr) {
			return true
		}
	} else {
		log.Println("！！！！选择错误！！！！！")
	}

	return false
}
