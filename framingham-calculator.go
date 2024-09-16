package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	//excel control library
	"github.com/xuri/excelize/v2"

	//GUI for users to select files
	"github.com/ncruces/zenity"
)

func calculateRisk(gender string, age, cholesterol, hdl, systolic, diastolic int, diabetes, smoking string) (int, float64, float64) {
	//define age points
	agePoints := map[string][]int{
		"female": {-9, -4, 0, 3, 6, 7, 8, 8, 8},
		"male":   {-1, 0, 1, 2, 3, 4, 5, 6, 7},
	}
	cholesterolPoints := map[string][]int{
		"female": {-2, 0, 1, 1, 3},
		"male":   {-3, 0, 1, 2, 3},
	}
	hdlPoints := map[string][]int{
		"female": {5, 2, 1, 0, -3},
		"male":   {2, 1, 0, 0, -3},
	}
	bpPoints := map[string][]int{
		"female": {-3, 0, 0, 2, 3},
		"male":   {0, 0, 1, 2, 3},
	}

	estimateRiskPoints := map[string][]int{
		"female": {0, 1, 2, 3, 5, 7, 8, 8, 8},
		"male":   {2, 3, 4, 4, 6, 7, 9, 11, 14},
	}

	//calculate age points
	ageIndex := (age - 30) / 5
	if ageIndex < 0 || ageIndex > 8 {
		log.Fatal("Invalid age for Framingham risk calculation.")
	}
	ageScore := agePoints[gender][ageIndex]

	//estimateRiskPoints
	estimateRiskIndex := (age - 30) / 5
	if estimateRiskIndex < 0 || estimateRiskIndex > 8 {
		log.Fatal("Invalid age for Framingham risk estimate range")
	}
	estimateRiskScore := estimateRiskPoints[gender][estimateRiskIndex]

	//calculate T-CHO points
	cholesterolIndex := 0
	if cholesterol >= 160 && cholesterol < 200 {
		cholesterolIndex = 1
	} else if cholesterol >= 200 && cholesterol < 240 {
		cholesterolIndex = 2
	} else if cholesterol >= 240 && cholesterol < 280 {
		cholesterolIndex = 3
	} else if cholesterol >= 280 {
		cholesterolIndex = 4
	}
	cholesterolScore := cholesterolPoints[gender][cholesterolIndex]

	//calculate HDL points
	hdlIndex := 0
	if hdl >= 35 && hdl < 45 {
		hdlIndex = 1
	} else if hdl >= 45 && hdl < 50 {
		hdlIndex = 2
	} else if hdl >= 50 && hdl < 60 {
		hdlIndex = 3
	} else if hdl >= 60 {
		hdlIndex = 4
	}
	hdlScore := hdlPoints[gender][hdlIndex]

	//calculate BP points
	bpIndex := 0
	if (systolic >= 120 && systolic < 130) || (diastolic >= 80 && diastolic < 85) {
		bpIndex = 1
	} else if (systolic >= 130 && systolic < 140) || (diastolic >= 85 && diastolic < 90) {
		bpIndex = 2
	} else if (systolic >= 140 && systolic < 160) || (diastolic >= 90 && diastolic < 100) {
		bpIndex = 3
	} else if systolic >= 160 || diastolic >= 100 {
		bpIndex = 4
	}
	bpScore := bpPoints[gender][bpIndex]

	//calculate diabetes points
	diabetesScore := 0
	if diabetes == "是" {
		if gender == "female" {
			diabetesScore = 4
		} else {
			diabetesScore = 2
		}
	}

	//calculate smoking points
	smokingScore := 0
	if smoking == "是" {
		smokingScore = 2
	}

	//total score
	totalScore := ageScore + cholesterolScore + hdlScore + bpScore + diabetesScore + smokingScore

	risk := float64(totalScore) * 0.01
	estimate := float64(estimateRiskScore) * 0.01

	return totalScore, risk, estimate
}

func toAlphaString(index int) string {
	result := ""
	for index >= 0 {
		result = string('A'+(index%26)) + result
		index = index/26 - 1
	}
	return result
}

func main() {
	//open a dialog for select files
	filePath, err := zenity.SelectFile()
	if err != nil {
		log.Fatal(err)
	}

	//open the excel file
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatal(err)
	}

	//get the existing sheet "工作表1"
	sheetName := "工作表1"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	//add new columns for risk calculation results
	for i, row := range rows {
		//skip the header row
		if i == 0 {
			f.SetCellValue(sheetName, fmt.Sprintf("K%d", i+1), "十年內發生缺血性心臟病的機率")
			f.SetCellValue(sheetName, fmt.Sprintf("L%d", i+1), "估計發生率")
			continue
		}

		//parse values from row
		if len(row) < 10 {
			log.Printf("Row %d does not have enough columns", i)
			continue
		}

		age, err := strconv.Atoi(row[3])
		if err != nil {
			log.Printf("Error parsing age in row %d: %v", i+1, err)
			continue
		}

		cholesterol, _ := strconv.Atoi(row[4])
		hdl, _ := strconv.Atoi(row[5])
		systolic, _ := strconv.Atoi(row[6])
		diastolic, _ := strconv.Atoi(row[7])
		gender := row[2]
		diabetes := row[8]
		smoking := row[9]

		// calculate risk
		_, risk, estimate := calculateRisk(gender, age, cholesterol, hdl, systolic, diastolic, diabetes, smoking)

		//write down the data to the existing sheet "工作表1"
		log.Printf("Writing data for row %d: risk=%.2f, estimate=%.2f", i, risk, estimate)
		f.SetCellValue(sheetName, fmt.Sprintf("K%d", i+1), fmt.Sprintf("%.2f", risk))
		f.SetCellValue(sheetName, fmt.Sprintf("L%d", i+1), fmt.Sprintf("%.2f", estimate))
	}

	//save the file
	savePath, err := zenity.SelectFileSave()
	if err != nil {
		log.Fatal(err)
	}

	// ensure the file has a .xlsx extension
	if !strings.HasSuffix(savePath, ".xlsx") {
		savePath += ".xlsx"
	}

	log.Printf("Saving file to: %s", savePath)
	if err := f.SaveAs(savePath); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Calculation complete.")
}
