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

func getfileXLSX() (listnama []string) {
	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}
	for _, file := range files {
		fn := file.Name()
		fn = fn[len(fn)-5:]

		if fn == ".xlsx" {
			listnama = append(listnama, file.Name())
		}
	}
	return listnama
}

func main() {
	pathlist := getfileXLSX()
	for _, path := range pathlist {
		//membuat database
		namadb := path[:len(path)-5] + ".db"
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
		for i := 0; i < banyakSheet; i++ {
			sheetName := f.GetSheetName(i)
			rows, err := f.Rows(sheetName)
			allrow, _ := f.GetRows(sheetName)
			totalbaris := len(allrow)
			if err != nil {
				fmt.Println(err, "copy kelima")
			}

			//membuat tabel database
			db, _ := sql.Open("sqlite3", "./"+namadb) // Open the created SQLite File
			execs := "CREATE TABLE " + sheetName + " (NO TEXT, KODE TEXT, URAIAN TEXT, SPESIFIKASI TEXT, SATUAN TEXT, HARGA_SATUAN TEXT, KETERANGAN TEXT, SUBJUDUL TEXT)"
			statement, err := db.Prepare(execs)
			if err != nil {
				fmt.Println("keempat" + err.Error())
			}
			statement.Exec()

			//next untuk melewati judul
			rows.Next()
			_, err = rows.Columns()
			if err != nil {
				fmt.Println(err, "keenam")
			}

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
				rows.Next()
				row, err := rows.Columns()
				if err != nil {
					fmt.Println(err, "ketujuh")
				}
				x := 0
				//memproses data perkolom pada setiap baris
				var data [7]string
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
				text := ""
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
					execs := "INSERT INTO " + sheetName + ` VALUES ("` + strconv.Itoa(noIncrement) + `",` + text + ")"
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
			db.Close()

		}
	}
	fmt.Println("Process complete!")

}
