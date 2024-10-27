package main

import (
	"fmt"
	"log"
	"net/http"
	"path/filepath"
	"strconv"
	"time"

	"github.com/gofiber/fiber/v2"
	"github.com/xuri/excelize/v2"
)

type User struct {
	No    int    `json:"no"`
	Name  string `json:"name"`
	Email string `json:"email"`
}

func generate(ctx *fiber.Ctx) error {
	path, err := generateExcell()
	if err != nil {
		return ctx.SendStatus(http.StatusInternalServerError)
	}

	return ctx.SendString(path)
}

func generateExcell() (string, error) {
	users := []User{
		{No: 1, Name: "FianGumilar", Email: "fiangumilar@gmail.com"},
		{No: 2, Name: "Harina Darmastuti", Email: "harina@gmail.com"},
	}

	// Create new file
	f := excelize.NewFile()
	defer f.Close()

	// Set new Sheet
	index, _ := f.NewSheet("Sheet1")
	f.SetActiveSheet(index)

	// Style for  border & fill 0
	sty0, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Style: 1, Color: "000000"},
			{Type: "top", Style: 1, Color: "000000"},
			{Type: "right", Style: 1, Color: "000000"},
			{Type: "bottom", Style: 1, Color: "000000"},
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"fdffd9"},
		},
	})

	// Styling for border & fill 1
	sty1, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Style: 1, Color: "000000"},
			{Type: "top", Style: 1, Color: "000000"},
			{Type: "right", Style: 1, Color: "000000"},
			{Type: "bottom", Style: 1, Color: "000000"},
		},
	})

	_ = f.SetCellValue("Sheet1", "A2", "No")
	_ = f.SetCellValue("Sheet1", "B2", "Name")
	_ = f.SetCellValue("Sheet1", "C2", "Email")

	_ = f.SetColWidth("Sheet1", "A", "A", 3.5)
	_ = f.SetColWidth("Sheet1", "B", "B", 20.85)
	_ = f.SetColWidth("Sheet1", "C", "C", 32.00)

	// Set Column Style
	_ = f.SetCellStyle("Sheet1", "A2", "C2", sty0)

	for idx, val := range users {
		colA := fmt.Sprintf("A%d", idx+3)
		colB := fmt.Sprintf("B%d", idx+3)
		colC := fmt.Sprintf("C%d", idx+3)

		// Set Col Style sty1
		_ = f.SetCellValue("Sheet1", colA, idx+1)
		_ = f.SetCellValue("Sheet1", colB, val.Name)
		_ = f.SetCellValue("Sheet1", colC, val.Email)

		// Set Style Col
		_ = f.SetCellStyle("Sheet1", colA, colC, sty1)
	}

	// Set save location
	location := "/Users/fiangumilar/Documents/belajar/go-generate-excell/storage"
	filename := strconv.FormatInt(time.Now().UnixNano(), 10) + ".xlsx"

	// Save As file
	if err := f.SaveAs(filepath.Join(location, filename)); err != nil {
		log.Printf("Error Save excel file")
	}

	return filename, nil
}

func main() {
	app := fiber.New()

	app.Get("/api/generate-data", generate)

	app.Static("/storage", "/Users/fiangumilar/Documents/belajar/go-generate-excell/storage")

	_ = app.Listen(":8991")
}
