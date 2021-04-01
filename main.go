package main

import (
  "fmt"
  "reflect"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type PoiData struct {
	Id                string `xlsx:"Id"`
	StationName       string `xlsx:"Station_Name" validate:"required"`
	StationName_en_BS string `xlsx:"StationName_en_BS"`
	StationName_zh_CN string `xlsx:"StationName_zh_CN"`
	StationName_zh_HK string `xlsx:"StationName_zh_HK"`
	StationName_zh_TW string `xlsx:"StationName_zh_TW"`
	StationName_ko_KR string `xlsx:"StationName_ko_KR"`
	StationName_th_TH string `xlsx:"StationName_th_TH"`
	StationName_vi_VN string `xlsx:"StationName_vi_VN"`
	StationName_id_ID string `xlsx:"StationName_id_ID"`
	StationName_ja_JP string `xlsx:"StationName_ja_JP"`
	StationName_en_US string `xlsx:"StationName_en_US"`
	StationName_en_AU string `xlsx:"StationName_en_AU"`
	StationName_en_NZ string `xlsx:"StationName_en_NZ"`
	StationName_en_GB string `xlsx:"StationName_en_GB"`
	StationName_en_SG string `xlsx:"StationName_en_SG"`
	StationName_en_IN string `xlsx:"StationName_en_IN"`
	StationName_en_CA string `xlsx:"StationName_en_CA"`
	StationName_en_HK string `xlsx:"StationName_en_HK"`
	StationName_en_PH string `xlsx:"StationName_en_PH"`
	StationName_en_MY string `xlsx:"StationName_en_MY"`
	StationName_fr_FR string `xlsx:"StationName_fr_FR"`
	StationName_es_ES string `xlsx:"StationName_es_ES"`
	StationName_de_DE string `xlsx:"StationName_de_DE"`
	StationName_it_IT string `xlsx:"StationName_it_IT"`
	StationName_ru_RU string `xlsx:"StationName_ru_RU"`

	Address                string `xlsx:"Address" validate:"required_with=Longitude Latitude"`
	Address_en_BS          string `xlsx:"Address_en_BS"`
	Address_zh_CN          string `xlsx:"Address_zh_CN"`
	Address_zh_HK          string `xlsx:"Address_zh_HK"`
	Address_zh_TW          string `xlsx:"Address_zh_TW"`
	Address_ko_KR          string `xlsx:"Address_ko_KR"`
	Address_th_TH          string `xlsx:"Address_th_TH"`
	Address_vi_VN          string `xlsx:"Address_vi_VN"`
	Address_id_ID          string `xlsx:"Address_id_ID"`
	Address_ja_JP          string `xlsx:"Address_ja_JP"`
	Address_en_US          string `xlsx:"Address_en_US"`
	Address_en_AU          string `xlsx:"Address_en_AU"`
	Address_en_NZ          string `xlsx:"Address_en_NZ"`
	Address_en_GB          string `xlsx:"Address_en_GB"`
	Address_en_SG          string `xlsx:"Address_en_SG"`
	Address_en_IN          string `xlsx:"Address_en_IN"`
	Address_en_CA          string `xlsx:"Address_en_CA"`
	Address_en_HK          string `xlsx:"Address_en_HK"`
	Address_en_PH          string `xlsx:"Address_en_PH"`
	Address_en_MY          string `xlsx:"Address_en_MY"`
	Address_fr_FR          string `xlsx:"Address_fr_FR"`
	Address_es_ES          string `xlsx:"Address_es_ES"`
	Address_de_DE          string `xlsx:"Address_de_DE"`
	Address_it_IT          string `xlsx:"Address_it_IT"`
	Address_ru_RU          string `xlsx:"Address_ru_RU"`
	Type                   string `xlsx:"Type" validate:"required"`
	ProductType            string `xlsx:"Product_Type"`
	AggregatorPositionCode string `xlsx:"Aggregator_Position_Code" validate:"required"`
	AggregatorName         string `xlsx:"Aggregator_Name" validate:"required"`
	Active                 int    `xlsx:"Active" validate:"oneof=0 1"`
	/**************poi 信息**************/
  Longitude float64 `xlsx:"Longitude" validate:"longitude"`
	//Latitude               float64 `xlsx:"Latitude" validate:"latitude"`
	Latitude float64 `xlsx:"Latitude"`
	//ProductType            string  `xlsx:"Product_Type" validate:"required"`
	//City                   string  `xlsx:"City" validate:"required_with=AreaID"`
	City   string `xlsx:"City"`
	AreaID int    `xlsx:"AreaID" validate:"numeric"`

	PlaceID        string `xlsx:"PlaceID"`
	PostCode       string `xlsx:"Post_Code"`
	PlatformAreaID int    `xlsx:"Platform_Area_ID"`
	GeoHash        string `xlsx:"Geo_Hash"`
  /**************poi 信息**************/
}



func (p PoiData) GetXLSXSheetName() string {
	return "Poi_Data"
}


func (p *PoiData) Excel(f *excelize.File, row int) {
	sheetName := p.GetXLSXSheetName()
	f.NewSheet(sheetName)
	fields := reflect.TypeOf(p).Elem()
	if err := f.SetCellValue(sheetName, Div(1)+fmt.Sprintf("%d", row), row-1); err != nil {
    fmt.Printf("excel生成错误", err)
  }
  valueOf := reflect.ValueOf(p)
  for i := 0; i < fields.NumField(); i++ {
	  eleName := fields.Field(i).Name
		rowVal := ""
		switch fields.Field(i).Type.Kind() {
		case reflect.String:
			rowVal = valueOf.Elem().FieldByName(eleName).String()
		case reflect.Float64, reflect.Float32:
			rowVal = fmt.Sprintf("%f", valueOf.Elem().FieldByName(eleName).Float())
		case reflect.Int, reflect.Int8, reflect.Int16,
			reflect.Int32, reflect.Int64, reflect.Uint8,
			reflect.Uint16, reflect.Uint32, reflect.Uint64:
			rowVal = fmt.Sprintf("%d", valueOf.Elem().FieldByName(eleName).Int())

		}
    err := f.SetCellValue(sheetName, Div(i+1)+fmt.Sprintf("%d", row), rowVal)
		if err != nil {
			fmt.Println("excel生成错误", err)
		}
	}
}


func Div(Num int) string {
	const alphabetNum = 26
	var (
		Str  string = ""
		k    int
		temp []int
	)
	//用来匹配的字符A-Z
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	if Num > alphabetNum {
		for {
			k = Num % alphabetNum
			if k == 0 {
				temp = append(temp, alphabetNum)
				k = alphabetNum
			} else {
				temp = append(temp, k)
			}
			Num = (Num - k) / alphabetNum
			if Num <= alphabetNum {
				temp = append(temp, Num)
				break
			}
		}
	} else {
		return Slice[Num]
	}

	for _, value := range temp {
		Str = Slice[value] + Str //因为数据切分后存储顺序是反的，所以Str要放在后面
	}
	return Str
}



// 生成表头
// 最多支持26列
// 通过xlsx tag 生成
func NewSheetANDTableHead(f *excelize.File, obj PoiData) {
	st := reflect.TypeOf(obj)
	if st.Kind() == reflect.Ptr {
		st = reflect.TypeOf(obj).Elem()
	}
	sheetName := obj.GetXLSXSheetName()
	f.NewSheet(sheetName)
	for i := 0; i < st.NumField(); i++ {
		rowName := Div(i) + "1"
		colName := st.Field(i).Tag.Get("xlsx")
		if err := f.SetCellValue(sheetName, rowName, colName); err != nil {
			fmt.Printf("excel 写入错误，%v", err)
		}
	}
}

func main() {
  outfile := excelize.NewFile()
  poiData := []PoiData{{
    StationName: "test",
    Longitude: 123.999,
    Latitude: 52.222,
    
  }}
  for key, poiItem := range poiData {
		if key == 0 {
		  NewSheetANDTableHead(outfile, poiItem)
		}
		fmt.Printf("poiItem: %v", poiItem)
		poiItem.Excel(outfile, (key + 2))
	}
  outfile.SaveAs("test.xlsx")
}