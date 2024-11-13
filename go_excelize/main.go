package main

import (
	"compress/gzip"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"strconv"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

type Record struct {
	ID               string  `json:"id"`
	MyString1        string  `json:"myString1"`
	MyDate1          string  `json:"myDate1"`
	MyDate2          string  `json:"myDate2"`
	Amount           float64 `json:"amount"`
	MyNumericString string `json:"myNumericString"`
	MyString2        string `json:"myString2"`
}

func getContent() ([]Record, error) {
	file, err := os.Open("../input.json.gzip")
	if err != nil {
		return nil, err
	}
	defer file.Close()

	gz, err := gzip.NewReader(file)
	if err != nil {
		return nil, err
	}
	defer gz.Close()

	contents, err := ioutil.ReadAll(gz)
	if err != nil {
		return nil, err
	}

	var records []Record
	err = json.Unmarshal(contents, &records)
	if err != nil {
		return nil, err
	}

	return records, nil
}

func toExcel(records []Record, nSheets int) error {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Set number formats
	decimalStyle, err := f.NewStyle(&excelize.Style{
		NumFmt: 3, // "0.000"
	})
	if err != nil {
		return err
	}

	dateStyle, err := f.NewStyle(&excelize.Style{
		NumFmt: 14, // "yyyy-mm-dd"
	})
	if err != nil {
		return err
	}

	var wg sync.WaitGroup
	errChan := make(chan error, nSheets)
	for i := 0; i < nSheets; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i+1)
		if i > 0 {
			f.NewSheet(sheetName)
		}
	}
	for i := 0; i < nSheets; i++ {
		wg.Add(1)
		go func(sheetIndex int) {
			defer wg.Done()

			sheetName := fmt.Sprintf("Sheet%d", sheetIndex+1)

			// Set column widths
			f.SetColWidth(sheetName, "A", "G", 22)

			// Create a new stream writer
			sw, err := f.NewStreamWriter(sheetName)
			if err != nil {
				errChan <- err
				return
			}

			// Write header row
			headerRow := []interface{}{"ID", "MyString1", "MyNumericString", "MyString2", "Amount", "MyDate2", "MyDate1"}
			if err := sw.SetRow("A1", headerRow); err != nil {
				errChan <- err
				return
			}

			for i, rec := range records {
				// fmt.Printf("Mystring %s \n", rec)
				row := i + 2 // Start from row 2 (row 1 is header)
				cells := make([]interface{}, 7)
				cells[0] = rec.ID
				cells[1] = rec.MyString1
				cells[2] = rec.MyNumericString
				cells[3] = rec.MyString2
				cells[4] = excelize.Cell{Value: rec.Amount, StyleID: decimalStyle}

				myDate2, _ := time.Parse("2006-01-02", rec.MyDate2)
				cells[5] = excelize.Cell{Value: myDate2, StyleID: dateStyle}

				myDate1, _ := time.Parse("2006-01-02", rec.MyDate1)
				cells[6] = excelize.Cell{Value: myDate1, StyleID: dateStyle}

				if err := sw.SetRow(fmt.Sprintf("A%d", row), cells); err != nil {
					errChan <- err
					return
				}
			}

			// Flush the stream
			if err := sw.Flush(); err != nil {
				errChan <- err
				return
			}
		}(i)
	}

	wg.Wait()
	close(errChan)

	for err := range errChan {
		if err != nil {
			return err
		}
	}

	return f.SaveAs("demo.xlsx")
}

func main() {
	nSheets := 1
	if nSheetsEnv := os.Getenv("N_SHEETS"); nSheetsEnv != "" {
		n, err := strconv.Atoi(nSheetsEnv)
		if err != nil || n < 1 || n > 9 {
			log.Fatalf("Invalid N_SHEETS value: %s. It should be between 1 and 9.", nSheetsEnv)
		}
		nSheets = n
	}

	start := time.Now()
	records, err := getContent()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Load Time %.2f seconds\n", time.Since(start).Seconds())

	start = time.Now()
	err = toExcel(records, nSheets)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Xlsx Write Time %.2f seconds\n", time.Since(start).Seconds())
}