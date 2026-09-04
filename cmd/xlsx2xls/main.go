//go:build windows

package main

import (
	_ "embed"
	"encoding/base64"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"unsafe"

	"github.com/dop251/goja"
)

//go:embed xlsx.full.min.js
var sheetJS string

const (
	OFN_EXPLORER        = 0x00080000
	OFN_FILEMUSTEXIST   = 0x00001000
	OFN_ALLOWMULTISELECT = 0x00000200
	MB_OK               = 0x00000000
	MB_ICONINFORMATION  = 0x00000040
	MB_ICONERROR        = 0x00000010
)

var (
	comdlg32              = syscall.NewLazyDLL("comdlg32.dll")
	user32                = syscall.NewLazyDLL("user32.dll")
	procGetOpenFileNameW  = comdlg32.NewProc("GetOpenFileNameW")
	procMessageBoxW       = user32.NewProc("MessageBoxW")
)

type OPENFILENAMEW struct {
	LStructSize       uint32
	HwndOwner         uintptr
	HInstance         uintptr
	LpstrFilter       *uint16
	LpstrCustomFilter *uint16
	NMaxCustFilter    uint32
	NFilterIndex      uint32
	LpstrFile         *uint16
	NMaxFile          uint32
	LpstrFileTitle    *uint16
	NMaxFileTitle     uint32
	LpstrInitialDir   *uint16
	LpstrTitle        *uint16
	Flags             uint32
	NFileOffset       uint16
	NFileExtension    uint16
	LpstrDefExt       *uint16
	LCustData         uintptr
	LpfnHook          uintptr
	LpTemplateName    *uint16
	PvReserved        uintptr
	DwReserved        uint32
	FlagsEx           uint32
}

func utf16Ptr(s string) *uint16 {
	p, _ := syscall.UTF16PtrFromString(s)
	return p
}

func messageBox(title, text string, icon uintptr) {
	t := utf16Ptr(title)
	m := utf16Ptr(text)
	procMessageBoxW.Call(0, uintptr(unsafe.Pointer(m)), uintptr(unsafe.Pointer(t)), MB_OK|icon)
}

func selectXlsxFiles() ([]string, error) {
	filter := []uint16{}
	for _, s := range []string{"Excel XLSX (*.xlsx)", "*.xlsx", "All Files (*.*)", "*.*", ""} {
		u, _ := syscall.UTF16FromString(s)
		filter = append(filter, u...)
	}
	buf := make([]uint16, 65536)
	title := utf16Ptr("選擇要轉換的 XLSX（可多選）")

	ofn := OPENFILENAMEW{
		LStructSize: uint32(unsafe.Sizeof(OPENFILENAMEW{})),
		LpstrFilter: &filter[0],
		NFilterIndex: 1,
		LpstrFile: &buf[0],
		NMaxFile: uint32(len(buf)),
		LpstrTitle: title,
		Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT,
	}
	r, _, _ := procGetOpenFileNameW.Call(uintptr(unsafe.Pointer(&ofn)))
	if r == 0 {
		return nil, errors.New("cancelled")
	}

	parts := []string{}
	start := 0
	for i, v := range buf {
		if v == 0 {
			if i == start {
				break
			}
			parts = append(parts, syscall.UTF16ToString(buf[start:i]))
			start = i + 1
		}
	}
	if len(parts) == 0 {
		return nil, errors.New("no file selected")
	}
	if len(parts) == 1 {
		return []string{parts[0]}, nil
	}

	dir := parts[0]
	out := make([]string, 0, len(parts)-1)
	for _, name := range parts[1:] {
		out = append(out, filepath.Join(dir, name))
	}
	return out, nil
}

func uniqueOutput(outDir, base string) string {
	p := filepath.Join(outDir, base+".xls")
	if _, err := os.Stat(p); os.IsNotExist(err) {
		return p
	}
	for i := 1; i < 10000; i++ {
		p = filepath.Join(outDir, fmt.Sprintf("%s_converted_%d.xls", base, i))
		if _, err := os.Stat(p); os.IsNotExist(err) {
			return p
		}
	}
	return filepath.Join(outDir, base+"_converted.xls")
}

func validateOLE(data []byte) bool {
	sig := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	if len(data) < len(sig) {
		return false
	}
	for i := range sig {
		if data[i] != sig[i] {
			return false
		}
	}
	return true
}

func convertXlsxToXls(input string) (string, error) {
	raw, err := os.ReadFile(input)
	if err != nil {
		return "", err
	}

	vm := goja.New()
	_ = vm.Set("window", vm.GlobalObject())
	_ = vm.Set("self", vm.GlobalObject())
	_ = vm.Set("global", vm.GlobalObject())

	if _, err := vm.RunString(sheetJS); err != nil {
		return "", fmt.Errorf("SheetJS 初始化失敗: %w", err)
	}

	b64 := base64.StdEncoding.EncodeToString(raw)
	_ = vm.Set("__INPUT_B64__", b64)

	script := `
	(function(){
		var wb = XLSX.read(__INPUT_B64__, {
			type: "base64",
			cellFormula: true,
			cellStyles: true,
			cellDates: true,
			cellNF: true,
			cellText: true
		});

		for (var i = 0; i < wb.SheetNames.length; ++i) {
			var ws = wb.Sheets[wb.SheetNames[i]];
			if (!ws || !ws["!ref"]) continue;
			var r = XLSX.utils.decode_range(ws["!ref"]);
			if (r.e.r >= 65536) throw new Error("XLS 97-2003 每個工作表最多 65536 列");
			if (r.e.c >= 256) throw new Error("XLS 97-2003 每個工作表最多 256 欄");
		}

		return XLSX.write(wb, {
			bookType: "biff8",
			type: "base64",
			bookSST: true,
			cellStyles: true
		});
	})()
	`

	v, err := vm.RunString(script)
	if err != nil {
		return "", fmt.Errorf("轉換失敗: %w", err)
	}

	outB64 := v.String()
	outBytes, err := base64.StdEncoding.DecodeString(outB64)
	if err != nil {
		return "", fmt.Errorf("輸出解碼失敗: %w", err)
	}
	if !validateOLE(outBytes) {
		return "", errors.New("輸出不是有效的 OLE/BIFF8 XLS")
	}

	srcDir := filepath.Dir(input)
	outDir := filepath.Join(srcDir, "XLS_轉換完成")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return "", err
	}
	base := strings.TrimSuffix(filepath.Base(input), filepath.Ext(input))
	dst := uniqueOutput(outDir, base)
	if err := os.WriteFile(dst, outBytes, 0644); err != nil {
		return "", err
	}
	return dst, nil
}

func main() {
	files, err := selectXlsxFiles()
	if err != nil {
		return
	}

	okCount := 0
	failCount := 0
	var failures []string

	for _, f := range files {
		if strings.ToLower(filepath.Ext(f)) != ".xlsx" {
			continue
		}
		if _, err := convertXlsxToXls(f); err != nil {
			failCount++
			failures = append(failures, filepath.Base(f)+"： "+err.Error())
		} else {
			okCount++
		}
	}

	msg := fmt.Sprintf("轉換完成\n\n成功：%d\n失敗：%d\n\n輸出資料夾：來源檔案旁的 XLS_轉換完成", okCount, failCount)
	if len(failures) > 0 {
		msg += "\n\n失敗項目：\n" + strings.Join(failures, "\n")
		messageBox("XLSX_to_XLS V5.0", msg, MB_ICONERROR)
	} else {
		messageBox("XLSX_to_XLS V5.0", msg, MB_ICONINFORMATION)
	}
}
