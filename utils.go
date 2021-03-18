package xlsx_utils

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	go_tool "github.com/maxiloEmmmm/go-tool"
	"io"
	"strings"
)

type Table struct {
	Cols []*TableCol
}

type lastRecord struct {
	Row   int
	Value interface{}
}

func (t *Table) XlsX(data interface{}, sheet string, w io.Writer) (err error) {
	return t.xlsX(data, sheet, w)
}

func (t *Table) XlsXDefaultSheet(data interface{}, w io.Writer) (err error) {
	return t.xlsX(data, "Sheet1", w)
}

func (t *Table) xlsX(data interface{}, sheet string, w io.Writer) (err error) {
	f := excelize.NewFile()
	fs := f.GetSheetName(f.NewSheet(sheet))
	if err = makeCol(f, fs, 1, 1, t.HeaderDepth(), t.Cols); err != nil {
		return
	}
	ids := t.getIndex()
	depth := t.HeaderDepth()
	v := go_tool.TryInterfacePtr(data)
	relationMap := make(map[string]*lastRecord, len(ids))

	for i := 0; i < v.Len(); i++ {
		for index, col := range ids {
			val, has := go_tool.Get(v.Index(i).Interface(), col.DataIndex)
			write := true
			if !has {
				val = ""
			} else if col.Relation {
				if lr, rHas := relationMap[col.DataIndex]; rHas {
					if lr.Value == val {
						write = false
					} else {
						// 如果不一样则尝试merge
						if i-1 != lr.Row {
							// MergeCell vcell行就不加1了 因为是合并到上一行为止 这里已经是新的一行了
							if err = f.MergeCell(fs, GetCol(lr.Row+1+depth, index+1), GetCol(i+depth, index+1)); err != nil {
								return
							}
						}

						lr.Value = val
						lr.Row = i
					}
				} else {
					relationMap[col.DataIndex] = &lastRecord{
						Row:   i,
						Value: val,
					}
				}
			}
			if write {
				if err = f.SetCellValue(fs, GetCol(i+1+depth, index+1), val); err != nil {
					return
				}
			}
		}
	}
	return f.Write(w)
}

func (t *Table) getIndex() []*TableCol {
	return getIndex(t.Cols)
}

func getIndex(cs []*TableCol) []*TableCol {
	ci := make([]*TableCol, 0)
	for _, c := range cs {
		if !c.HasChild() {
			ci = append(ci, c)
		} else {
			ci = append(ci, getIndex(c.Children)...)
		}
	}
	return ci
}

func (t *Table) HeaderDepth() int {
	max := 1
	for _, col := range t.Cols {
		if cd := colDepth(col); cd > max {
			max = cd
		}
	}
	return max
}

const MaxCol = 10

func GetCol(row int, col int) string {
	colSet := make([]int, MaxCol)
	for i := 0; i < col; i++ {
		for j := 0; j < MaxCol; j++ {
			if colSet[j] == 26 {
				colSet[j] = 1
			} else {
				colSet[j]++
				break
			}
		}
	}
	builder := new(strings.Builder)
	for i := len(colSet) - 1; i >= 0; i-- {
		if colSet[i] > 0 {
			builder.WriteByte('A' + byte(colSet[i]-1))
		}
	}
	builder.WriteString(fmt.Sprintf("%d", row))
	return builder.String()
}

func makeCol(f *excelize.File, sheet string, colStart int, rowStart int, depth int, cols []*TableCol) (err error) {
	for ci := 0; ci < len(cols); ci++ {
		col := cols[ci]
		colNumber := ci + colStart
		if err = f.SetCellValue(sheet, GetCol(rowStart, colNumber), col.Title); err != nil {
			return
		}

		if col.HasChild() {
			// 合并多余列
			cw := col.HeaderColWidth() - 1
			target := GetCol(rowStart, colNumber+cw)
			if err = f.MergeCell(sheet, GetCol(rowStart, colNumber), target); err != nil {
				return
			}
			if err = makeCol(f, sheet, colNumber, rowStart+1, depth-1, col.Children); err != nil {
				return
			}
			// 因为包含子列 这里跨度要累加子列多出来的宽度
			colStart += cw
		} else if depth > 1 {
			// 合并多余行
			if err = f.MergeCell(sheet, GetCol(rowStart, colNumber), GetCol(rowStart+depth-1, colNumber)); err != nil {
				return
			}
		}
	}
	return
}

type TableCol struct {
	Title     string `json:"title"`
	DataIndex string `json:"dataIndex,omitempty"`
	// 提供行合并
	Relation bool
	Children []*TableCol `json:"children,omitempty"`
}

func (tc *TableCol) HasChild() bool {
	return len(tc.Children) > 0
}

func (tc *TableCol) HeaderColWidth() int {
	return colWidth(tc)
}

func colWidth(col *TableCol) int {
	d := 1
	if col.HasChild() {
		d = 0
		for _, cc := range col.Children {
			d += colWidth(cc)
		}
	}
	return d
}

func colDepth(col *TableCol) int {
	d := 1
	if col.HasChild() {
		max := 1
		for _, cc := range col.Children {
			if childDepth := colDepth(cc); childDepth > max {
				max = childDepth
			}
		}
		d += max
	}
	return d
}
