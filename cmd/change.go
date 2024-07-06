/*
Copyright © 2024 NAME HERE <EMAIL ADDRESS>
*/
package cmd

import (
	"archive/zip"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"text/template"
	"time"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

type Task struct {
	Name string `json:"name"`
	Hour int    `json:"hour"`
}

type Settings struct {
	Name      string `json:"name"`
	HourlyPay int    `json:"hourlyPay"`
	Year      int    `json:"year"`
	Month     int    `json:"month"`
	Tasks     []Task `json:"tasks"`
}

func loadSetting() (Settings, error) {
	var settings Settings
	jsonFile, err := os.Open("settings.json")
	if err != nil {
		log.Println(err)
		return settings, err
	}
	defer jsonFile.Close()
	byteValue, err := io.ReadAll(jsonFile)
	if err != nil {
		log.Println(err)
		return settings, err
	}
	err = json.Unmarshal(byteValue, &settings)
	if err != nil {
		log.Println(err)
		return settings, err
	}
	return settings, nil
}

func getEndOfMonth(year int, month int) time.Time {
	return time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.Local).AddDate(0, 1, -1)
}

func changeExcel(settings Settings) {
	const SHEET_NAME = "請求書"
	f, err := excelize.OpenFile("data/template/請求書.xlsx")
	if err != nil {
		log.Println(err)
		return
	}
	defer f.Close()

	setCellValue := func(cell string, value interface{}) {
		err = f.SetCellValue(SHEET_NAME, cell, value)
		if err != nil {
			log.Println(err)
		}
	}
	year := 2024
	month := 12
	setCellValue("O2", getEndOfMonth(year, month).Format("2006年1月2日"))
	setCellValue("D10", getEndOfMonth(year, month).Format("2006年1月2日"))

	// 案件
	for i, task := range settings.Tasks {
		row := 20 + i
		setCellValue(fmt.Sprintf("A%d", row), i+1)
		setCellValue(fmt.Sprintf("B%d", row), fmt.Sprintf("%s (%dh)", task.Name, task.Hour))
		setCellValue(fmt.Sprintf("J%d", row), 1)
		setCellValue(fmt.Sprintf("L%d", row), task.Hour*settings.HourlyPay)
	}

	err = f.UpdateLinkedValue()
	if err != nil {
		log.Println(err)
	}
	err = f.SaveAs(fmt.Sprintf("data/請求書_%s_%d年%d月_暫定版.xlsx", settings.Name, settings.Year, settings.Month))
	if err != nil {
		log.Println(err)
	}
}

func changeWord(settings Settings) {
	// xml ファイル作成
	{
		doc, err := template.ParseFiles("data/template/document.xml")
		if err != nil {
			log.Println(err)
			return
		}
		data := struct {
			Settings
			HourSum int
		}{
			Settings: settings,
			HourSum:  0,
		}
		for _, task := range settings.Tasks {
			data.HourSum += task.Hour
		}
		dest, err := os.Create("data/template/docx/word/document.xml") // 出力先
		if err != nil {
			log.Println(err)
			return
		}
		defer dest.Close()
		err = doc.Execute(dest, data)
		if err != nil {
			log.Println(err)
			return
		}
	}

	// docx ファイル作成
	{
		docx, err := os.Create(fmt.Sprintf("data/作業報告書_%s_%d年%d月_暫定版.docx", settings.Name, settings.Year, settings.Month))
		if err != nil {
			log.Println(err)
			return
		}
		defer docx.Close()
		writer := zip.NewWriter(docx)
		defer writer.Close()

		root := "data/template/docx"
		err = filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
			if err != nil {
				return err
			}
			if info.IsDir() {
				return nil
			}

			relPath, err := filepath.Rel(root, path)
			if err != nil {
				return err
			}

			w, err := writer.Create(relPath)
			if err != nil {
				return err
			}

			file, err := os.Open(path)
			if err != nil {
				return err
			}
			defer file.Close()

			_, err = io.Copy(w, file)
			if err != nil {
				return err
			}

			return nil
		})
		if err != nil {
			log.Println(err)
			return
		}
	}
}

// changeCmd represents the change command
var changeCmd = &cobra.Command{
	Use:   "change",
	Short: "A brief description of your command",
	Long: `A longer description that spans multiple lines and likely contains examples
and usage of using your command. For example:

Cobra is a CLI library for Go that empowers applications.
This application is a tool to generate the needed files
to quickly create a Cobra application.`,
	Run: func(cmd *cobra.Command, args []string) {
		settings, err := loadSetting()
		if err != nil {
			return
		}
		changeExcel(settings)
		changeWord(settings)
	},
}

func init() {
	rootCmd.AddCommand(changeCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// changeCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// changeCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}

// https://stkyotouac-my.sharepoint.com/:f:/g/personal/omori_tomohito_58z_st_kyoto-u_ac_jp/ErD8IGSX4rFAmtA3oj5LThYB_62axxvd2WWhvMBQJVSHVw
