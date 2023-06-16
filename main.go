package main

import (
	"fmt"
	"sync"
	"unsafe"

	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
)

/*
#include <stdlib.h>
typedef struct string_arr {
	char **arr;
	int str_size;
} string_arr;

typedef struct string_arr2 {
	char **arr;
	int *row_sizes;
	int rows;
} string_arr2;

typedef struct cus_obj {
	void *obj;
	uint32 typ;
} cus_obj
*/

import "C"

var files map[uint32]*excelize.File
var once sync.Once

func getFile(uid uint32) *excelize.File {
	return files[uid]
}

func setFile(uid uint32, file *excelize.File) {
	once.Do(func() {
		files = make(map[uint32]*excelize.File)
	})
	files[uid] = file
}

//export printStr
func printStr(str *C.char) {
	ss := C.GoString(str)
	fmt.Println(ss)
}

//export newFile
func newFile() uint32 {
	file := excelize.NewFile()
	uid := fuuid()
	setFile(uid, file)
	return uid
}

//export openFile
func openFile(filePath *C.char) uint32 {
	fmt.Println(filePath)
	path := C.GoString(filePath)
	fmt.Println(path)
	file, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		return 0
	}
	uid := fuuid()
	setFile(uid, file)
	// defer C.free(unsafe.Pointer(filePath))
	return uid
}

//export getSheetList
func getSheetList(fileId uint32) *C.struct_string_arr {
	f := getFile(fileId)
	names := f.GetSheetList()
	cArray := C.malloc(C.size_t(len(names)) * C.size_t(unsafe.Sizeof(uintptr(0))))

	a := (*[1<<30 - 1]*C.char)(cArray)

	for i, name := range names {
		a[i] = C.CString(name)
	}
	result := (*C.struct_string_arr)(C.malloc(C.size_t(unsafe.Sizeof(C.struct_string_arr{}))))
	result.arr = (**C.char)(cArray)
	result.str_size = C.int(len(names))
	return result
}

//export getSheetName
func getSheetName(fileId uint32, index int) *C.char {
	f := getFile(fileId)
	name := f.GetSheetName(index)
	return C.CString(name)
}

//export closeFile
func closeFile(fileId uint32) {
	f := getFile(fileId)
	f.Close()
	delete(files, fileId)
}

//export setSheetVisible
func setSheetVisible(fileId uint32, sheetNameC *C.char, visible C.int) {
	f := getFile(fileId)
	sheetName := C.GoString(sheetNameC)
	defer C.free(unsafe.Pointer(sheetNameC))
	f.SetSheetVisible(sheetName, int(visible) == 1)
}

//export deleteSheet
func deleteSheet(fileId uint32, sheetNameC *C.char) {
	f := getFile(fileId)
	sheetName := C.GoString(sheetNameC)
	defer C.free(unsafe.Pointer(sheetNameC))
	f.DeleteSheet(sheetName)
}

//export setCellValue
func setCellValue(fileId uint32, sheetNameC *C.char, row C.int, col C.int, value unsafe.Pointer, typ int) {
	f := getFile(fileId)
	sheetName := C.GoString(sheetNameC)
	// defer C.free(unsafe.Pointer(sheetNameC))
	cell, _ := excelize.CoordinatesToCellName(int(col), int(row))
	switch int(typ) {
	case 1: //int
		ptr := (*C.int)(value)
		f.SetCellInt(sheetName, cell, int(*ptr))
	case 2: //float
		ptr := (*C.float)(value)
		f.SetCellFloat(sheetName, cell, float64(*ptr), 4, 32)
	case 3: //string
		val := (*C.char)(value)
		f.SetCellStr(sheetName, cell, C.GoString(val))
	}
}

//export getCellValue
func getCellValue(fileId uint32, sheetNameC *C.char, row C.int, col C.int) *C.char {
	f := getFile(fileId)
	sheetName := C.GoString(sheetNameC)
	fmt.Println(fileId)
	cell, _ := excelize.CoordinatesToCellName(int(col), int(row))
	// defer C.free(unsafe.Pointer(sheetNameC))
	val, _ := f.GetCellValue(sheetName, cell)

	str := C.CString(val)
	// defer C.free(unsafe.Pointer(str))
	return str
}

//export getRows
func getRows(fileId uint32, sheetNameC *C.char) *C.struct_string_arr2 {
	f := getFile(fileId)
	sheetName := C.GoString(sheetNameC)
	rows, _ := f.GetRows(sheetName)
	// defer C.free(unsafe.Pointer(sheetNameC))
	var allCount int
	for _, row := range rows {
		allCount += len(row)
	}
	cArray := C.malloc(C.size_t(allCount) * C.size_t(unsafe.Sizeof(uintptr(0))))
	rowSizes := (*C.int)(C.malloc(C.size_t(len(rows)) * C.size_t(unsafe.Sizeof(C.int(0)))))
	a := (*[1<<30 - 1]*C.char)(cArray)
	// b := (*[1<<30 - 1]C.int)(rowSizes)
	i := 0
	for rowIndex, row := range rows {
		(*[1<<30 - 1]C.int)(unsafe.Pointer(rowSizes))[rowIndex] = C.int(len(row))
		for _, str := range row {
			a[i] = C.CString(str)
			i += 1
		}
	}
	result := (*C.struct_string_arr2)(C.malloc(C.size_t(unsafe.Sizeof(C.struct_string_arr2{}))))
	result.arr = (**C.char)(cArray)
	result.rows = C.int(len(rows))
	result.row_sizes = (*C.int)(rowSizes)
	return result
}

func putRows(fileId uint32, sheetNameC *C.char, *C.struct_string_arr) {
	f := getFile(fileId)
	sheetName := C.GoString(sheetNameC)

}

//export save
func save(fileId uint32) {
	f := getFile(fileId)
	f.Save()
}

//export saveAs
func saveAs(fileId uint32, path *C.char) {
	f := getFile(fileId)
	strPath := C.GoString(path)
	f.SaveAs(strPath)
}

// void freeStringArr2(string_arr2 *ptr) {
// 	int index = 0;
// 	for (int i = 0; i < ptr->rows; i ++) {
// 		for(int j = 0 ;j < (ptr -> row_sizes)[i]; j++) {
// 			free(ptr->arr[index]);
// 			index++;
// 		}
// 	}
// 	free(ptr -> arr);
// 	free(ptr -> row_sizes);
// 	free(ptr);
// }

//export freeStringArr2
func freeStringArr2(ptr *C.struct_string_arr2) {
	var index int
	arrPtrs := (*[1<<30 - 1]*C.char)(unsafe.Pointer(ptr.arr))
	for i := 0; i < int(ptr.rows); i++ {
		for j := 0; j < int((*[1<<30 - 1]C.int)(unsafe.Pointer(ptr.row_sizes))[i]); j++ {
			C.free(unsafe.Pointer(arrPtrs[index]))
			index += 1
		}
	}
	C.free(unsafe.Pointer(ptr.arr))
	C.free(unsafe.Pointer(ptr.row_sizes))
	C.free(unsafe.Pointer(ptr))
}

func fuuid() uint32 {
	return uuid.New().ID()
}

func main() {
}
