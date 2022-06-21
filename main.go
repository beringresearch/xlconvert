package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	sheet := flag.String("s", "Sheet1", "optional workbook Sheet Name")
	flag.Parse()

	if len(os.Args) < 2 {
		fmt.Println(`Convert XLSX files to CSV

Usage:
  xlconvert INFILE OUTFILE
  
Flags:
  -s Sheet name (default: Sheet1)`)
		os.Exit(0)
	}

	if len(os.Args) < 3 {
		log.Fatal("missing output file name")
	}

	fname := os.Args[1]
	outname := os.Args[2]

	f, err := excelize.OpenFile(fname)
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(*sheet)
	if err != nil {
		log.Fatal(err)
	}

	file, err := os.Create(outname)
	if err != nil {
		log.Fatal("unable to create file: ", err.Error())
	}

	writer := csv.NewWriter(file)
	defer writer.Flush()

	for _, value := range rows {
		err := writer.Write(value)
		if err != nil {
			log.Fatal("unable to write file: ", err.Error())
		}
	}
}
