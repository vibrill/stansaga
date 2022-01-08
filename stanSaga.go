package main

import (
	"database/sql"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"strconv"
	"strings"

	_ "github.com/mattn/go-sqlite3"
	"github.com/schollz/progressbar/v3"
	"github.com/xuri/excelize/v2"
)

func CreateDB(namadb string) {
	//createdb
	os.Remove("./" + namadb)              // I delete the file to avoid duplicated records.
	file, err := os.Create("./" + namadb) // Create SQLite file
	if err != nil {
		fmt.Println("ketiga" + err.Error())
	}
	file.Close()
}

func getfileXLSX() (listnama []string, detektor []string) {
	var listxls []string
	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}
	for _, file := range files {
		fn := file.Name()
		if len(fn) > 5 {
			x := len(fn) - 5
			if fn[x:] == ".xlsx" {
				listxls = append(listxls, file.Name())
			}
		}
	}
	for _, file := range listxls {
		f, err := excelize.OpenFile(file)
		if err == nil {
			deteksi := "MODAL.db"
			for i := 0; i < len(f.GetSheetList()); i++ {
				namaSheet := f.GetSheetName(i)
				if namaSheet == "SSH1" {
					deteksi = "BARJAS.db"
					break
				}
				if namaSheet == "KODE BELANJA" {
					deteksi = "KODEREK.db"
					break
				}
			}
			listnama = append(listnama, file)
			detektor = append(detektor, deteksi)
		}
	}
	return listnama, detektor
}

func main() {
	pathlist, detektor := getfileXLSX()
	for i := 0; i < len(pathlist); i++ {
		path := pathlist[i]
		namadb := detektor[i]
		//membuat database
		fmt.Println("membuat database", namadb)
		CreateDB(namadb)

		//excelize membuka file excel
		f, err := excelize.OpenFile(path)
		if err == nil {
			fmt.Println("Unmerge File", path)
		}
		banyakSheet := f.SheetCount

		for i := 0; i < banyakSheet; i++ {
			f.SetActiveSheet(i)
			//mendapatkan nama sheet
			sheetName := f.GetSheetName(i)
			//unmerge
			merge_cells, _ := f.GetMergeCells(sheetName)
			for i := range merge_cells {
				axis_x := strings.Split(merge_cells[i][0], ":")[0]
				axis_y := strings.Split(merge_cells[i][0], ":")[1]
				fmt.Printf("index=%d\tref=%s\tval=%s\t\tUnmergeCell(\"%s\",\"%s\",\"%s\")\t", i, merge_cells[i][0], merge_cells[i][1], sheetName, axis_x, axis_y)
				err = f.UnmergeCell(sheetName, axis_x, axis_y)
				if err != nil {
					fmt.Printf("UnmergeCell fail,%v\n", err)
					return
				} else {
					fmt.Printf("OK\n")
				}
			}
		}
		fmt.Println("unmerge done")
		f.Save()

		f, _ = excelize.OpenFile(path)
		banyakSheet = f.SheetCount
		//println("ini adalah banyak sheet:", banyakSheet)
		for i := 0; i < banyakSheet; i++ {
			sheetName := f.GetSheetName(i)
			//println("ini adalah nama sheet: ", sheetName)
			rows, err := f.Rows(sheetName)
			allrow, _ := f.GetRows(sheetName)
			totalbaris := len(allrow)
			//print("ini total baris :", totalbaris, "\n")
			if err != nil {
				fmt.Println(err, "copy kelima")
			}

			//membuat tabel database
			db, _ := sql.Open("sqlite3", "./"+namadb) // Open the created SQLite File
			execs := "CREATE TABLE " + strings.ReplaceAll(sheetName, " ", "_") + " (NO TEXT, KODE TEXT, URAIAN TEXT, SPESIFIKASI TEXT, SATUAN TEXT, HARGA_SATUAN TEXT, KETERANGAN TEXT, SUBJUDUL TEXT)"
			if namadb == "KODEREK.db" {
				execs = "CREATE TABLE " + strings.ReplaceAll(sheetName, " ", "_") + " (NO TEXT, JENIS_BELANJA TEXT, KODE TEXT, URAIAN TEXT, KETERANGAN TEXT)"
			}
			statement, err := db.Prepare(execs)
			if err != nil {
				fmt.Println("keempat" + err.Error())
			}
			statement.Exec()

			//next untuk melewati judul
			rows.Next()

			//mendapatkan data perbaris meletakkannnya pada list data
			fmt.Println("Generate Table", sheetName, ", please wait!")
			bariskosong := 0
			subjudul := "empty"
			//25780
			noIncrement := 0
			bar := progressbar.Default(int64(totalbaris))
			for i := 0; i < totalbaris; i++ {
				bar.Add(1)
				if bariskosong == 3 {
					break
				}
				row, err := rows.Columns()
				//println("ini adalah row : ", row)
				if err != nil {
					fmt.Println(err, "ketujuh")
				}

				//memproses data perkolom pada setiap baris

				var data [7]string
				var datakorek [4]string
				text := ""
				x := 0
				if namadb == "KODEREK.db" {
					rows.Next()
					for _, colCell := range row {
						//fmt.Println("ini adalah collcel: ", colCell)
						datakorek[x] = strings.ReplaceAll(colCell, `"`, "")
						x++
					}
					//deteksi kolom kosong
					kolomkosong := 0
					for _, item := range datakorek {
						if item == "" {
							kolomkosong += 1
						}
					}

					//jika kolom kosong == 4 baris kosong +1
					if kolomkosong == 4 {
						bariskosong += 1
					}
					if kolomkosong != 4 {
						bariskosong = 0
					}
					//fmt.Println("ini adalah datakorek : ", datakorek)
					for _, item := range datakorek {
						text = text + `"` + item + `",`

					}

					//fmt.Println("ini teks yang akan ditulis: ", text)
					text = text[:len(text)-1]
					noIncrement += 1
					if bariskosong != 1 {
						execs := "INSERT INTO " + strings.ReplaceAll(sheetName, " ", "_") + ` VALUES ("` + strconv.Itoa(noIncrement) + `",` + text + ")"
						statement, err := db.Prepare(execs) // Prepare statement.
						if err != nil {
							fmt.Println(err.Error())
						}
						_, err = statement.Exec()
						if err != nil {
							fmt.Println(err.Error())
						}
					}

				} else {
					rows.Next()
					for _, colCell := range row {
						data[x] = strings.ReplaceAll(colCell, `"`, "")
						x++
					}

					//deteksi kolom kosong
					kolomkosong := 0
					for _, item := range data {
						if item == "" {
							kolomkosong += 1
						}
					}

					//jika kolom kosong == 7 baris kosong +1
					if kolomkosong == 7 {
						bariskosong += 1
					}
					if kolomkosong != 7 {
						bariskosong = 0
					}

					//mendapatkan value cell per baris
					for _, item := range data { //sepadan dengan "for item in data:" pada python
						text = text + `"` + item + `",`
					}
					text = text[:len(text)-1] //omit last coma

					//mengubah indikatorsubjudul
					if kolomkosong == 6 {
						subjudul = data[0]
					}

					//memasukkan keterangan subjudul jika bukan subjudul dan baris kosong
					if kolomkosong != 6 && kolomkosong != 7 {
						text = text + `"` + subjudul + `"`
					} else {
						text = text + `"empty"`
					}

					//insert data
					if kolomkosong != 7 {
						noIncrement += 1
						execs := "INSERT INTO " + strings.ReplaceAll(sheetName, " ", "_") + ` VALUES ("` + strconv.Itoa(noIncrement) + `",` + text + ")"
						statement, err := db.Prepare(execs) // Prepare statement.
						if err != nil {
							fmt.Println(err.Error())
						}
						_, err = statement.Exec()
						if err != nil {
							fmt.Println(err.Error())
						}
					}
				}

			}
			db.Close()

		}
	}
	fmt.Println("Process complete!")

}
