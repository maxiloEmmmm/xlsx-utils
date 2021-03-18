package xlsx_utils

import (
	"github.com/stretchr/testify/require"
	"os"
	"testing"
)

// 根据数字获取单元格号码
func TestGetCol(t *testing.T) {
	require.Equal(t, GetCol(1, 1), "A1")
	require.Equal(t, GetCol(1, 26), "Z1")
	require.Equal(t, GetCol(1, 27), "AA1")
	require.Equal(t, GetCol(1, 28), "AB1")
	require.Equal(t, GetCol(1, 53), "BA1")
}

// 获取列的实际占用宽度
func TestColWidth(t *testing.T) {
	col := &TableCol{
		Title: "a",
		Children: []*TableCol{
			{Title: "b", DataIndex: "ab", Children: []*TableCol{
				{Title: "b", DataIndex: "acb"},
				{Title: "c", DataIndex: "acc"},
			}},
			{Title: "c", Children: []*TableCol{
				{Title: "b", DataIndex: "ab"},
				{Title: "b", DataIndex: "ab"},
				{Title: "b", DataIndex: "ab"},
			}},
		},
	}
	require.Equal(t, col.HeaderColWidth(), 5)
}

// 获取列的实际占用高度
func TestHeaderDepth(t *testing.T) {
	tab := &Table{Cols: []*TableCol{
		{Title: "a", Children: []*TableCol{
			{Title: "b", DataIndex: "ab"},
			{Title: "c", DataIndex: "ac"},
		}},
	}}
	require.Equal(t, tab.HeaderDepth(), 2)
}

const TestFile = "./test.xlsx"

// 获取导出
func TestSave(t *testing.T) {
	tab := &Table{Cols: []*TableCol{
		{Title: "a", Children: []*TableCol{
			{Title: "b", DataIndex: "ab"},
			{Title: "c", Children: []*TableCol{
				{Title: "b", DataIndex: "acb"},
				{Title: "c", DataIndex: "acc"},
			}},
		}},
	}}

	f, err := os.Create(TestFile)
	require.Nil(t, err)

	err = tab.XlsXDefaultSheet([]interface{}{
		map[string]interface{}{
			"ab":  1,
			"acb": "2",
			"acc": "3",
		},
	}, f)
	require.Nil(t, err)

	require.Nil(t, f.Close())
	require.Nil(t, os.Remove(TestFile))
}

// 获取map路径数据
func TestPathSave(t *testing.T) {
	tab := &Table{Cols: []*TableCol{
		{Title: "a", Children: []*TableCol{
			{Title: "b", DataIndex: "ab.b"},
			{Title: "c", Children: []*TableCol{
				{Title: "b", DataIndex: "acb"},
				{Title: "c", DataIndex: "acc"},
			}},
		}},
	}}

	f, err := os.Create(TestFile)
	require.Nil(t, err)

	err = tab.XlsXDefaultSheet([]interface{}{
		map[string]interface{}{
			"ab": map[string]interface{}{
				"b": 1,
			},
			"acb": "2",
			"acc": "3",
		},
	}, f)
	require.Nil(t, err)

	require.Nil(t, f.Close())
	require.Nil(t, os.Remove(TestFile))
}

type TestStructPath struct {
	Ab *TestStructPathChild
}

type TestStructPathChild struct {
	A string
	C string
}

// 获取结构路径数据
func TestStructPathSave(t *testing.T) {
	tab := &Table{Cols: []*TableCol{
		{Title: "a", Children: []*TableCol{
			{Title: "b", DataIndex: "Ab.A"},
			{Title: "b", DataIndex: "Ab.C"},
		}},
	}}

	f, err := os.Create(TestFile)
	require.Nil(t, err)

	err = tab.XlsXDefaultSheet([]interface{}{
		&TestStructPath{
			Ab: &TestStructPathChild{
				A: "1",
				C: "2",
			},
		},
		&TestStructPath{
			Ab: &TestStructPathChild{
				A: "3",
				C: "4",
			},
		},
	}, f)
	require.Nil(t, err)

	require.Nil(t, f.Close())
	require.Nil(t, os.Remove(TestFile))
}

// 获取合并导出
func TestRowMergeSave(t *testing.T) {
	tab := &Table{Cols: []*TableCol{
		{Title: "a", Children: []*TableCol{
			// 为ab列提供行合并
			{Title: "b", DataIndex: "ab", Relation: true},
			{Title: "c", Children: []*TableCol{
				{Title: "b", DataIndex: "acb"},
				{Title: "c", DataIndex: "acc"},
			}},
		}},
	}}

	f, err := os.Create(TestFile)
	require.Nil(t, err)

	err = tab.XlsXDefaultSheet([]interface{}{
		map[string]interface{}{
			"ab":  1,
			"acb": "2",
			"acc": "3",
		},
		// ab列的2和3行会合并
		map[string]interface{}{
			"ab":  2,
			"acb": "2",
			"acc": "3",
		},
		map[string]interface{}{
			"ab":  2,
			"acb": "2",
			"acc": "3",
		},
		map[string]interface{}{
			"ab":  3,
			"acb": "2",
			"acc": "3",
		},
	}, f)
	require.Nil(t, err)
	require.Nil(t, f.Close())
	require.Nil(t, os.Remove(TestFile))
}
