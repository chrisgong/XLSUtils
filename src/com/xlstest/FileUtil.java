package com.xlstest;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class FileUtil {

	/**
	 * 解析xls文件
	 * 
	 * @author chris
	 * @param path
	 *            xls文件的路径
	 * */
	public static String readXLS(String path) {
		String str = "解析文件出现问题";
		try {
			Workbook workbook = null;
			workbook = Workbook.getWorkbook(new File(path));
			Sheet sheet = workbook.getSheet(0);
			Cell cell = null;
			int columnCount = sheet.getColumns();
			int rowCount = sheet.getRows();
			for (int i = 0; i < rowCount; i++) {
				for (int j = 0; j < columnCount; j++) {
					cell = sheet.getCell(j, i);
					String temp2 = "";
					if (cell.getType() == CellType.NUMBER) {
						temp2 = ((NumberCell) cell).getValue() + "";
					} else if (cell.getType() == CellType.DATE) {
						temp2 = "" + ((DateCell) cell).getDate();
					} else {
						temp2 = "" + cell.getContents();
					}
					str = str + "  " + temp2;
				}
				str += "\n";
			}
			workbook.close();
		} catch (Exception e) {
			return str;
		}
		return str;
	}

	/**
	 * 导出xls文件
	 * 
	 * @author chris
	 * 
	 * @param dataList
	 *            文件数据源
	 * */
	public static void exportXLS(ArrayList<TestModel> dataList, String fileName) {
		WritableWorkbook wwb = null;
		try {
			// 首先要使用Workbook类的工厂方法创建一个可写入的工作薄(Workbook)对象
			wwb = Workbook.createWorkbook(new File(fileName + ".xls"));
		} catch (IOException e) {
			e.printStackTrace();
		}
		if (wwb != null) {
			// 创建一个可写入的工作表
			// Workbook的createSheet方法有两个参数，第一个是工作表的名称，第二个是工作表在工作薄中的位置
			WritableSheet ws = wwb.createSheet("工作表名称", 0);

			// 下面开始添加单元格
			String[] topic = { "id", "测量时间", "目标温度底数", "目标温度", "环境温度底数", "环境温度" };
			for (int i = 0; i < topic.length; i++) {
				Label labelC = new Label(i, 0, topic[i]);
				try {
					// 将生成的单元格添加到工作表中
					ws.addCell(labelC);
				} catch (RowsExceededException e) {
					e.printStackTrace();
				} catch (WriteException e) {
					e.printStackTrace();
				}
			}
			TestModel model;
			ArrayList<String> li;
			for (int i = 0; i < dataList.size(); i++) {
				model = dataList.get(i);
				li = new ArrayList<String>();
				li.add(model.getId());
				li.add(model.getDate());
				li.add(model.getTarget_temp_base());
				li.add(model.getTarget_temp());
				li.add(model.getEnvironmental_temp_base());
				li.add(model.getEnvironmental_temp());
				int k = 0;
				for (String l : li) {
					Label labelC = new Label(k, i + 1, l);
					k++;
					try {
						// 将生成的单元格添加到工作表中
						ws.addCell(labelC);
					} catch (RowsExceededException e) {
						e.printStackTrace();
					} catch (WriteException e) {
						e.printStackTrace();
					}
				}
				li = null;
			}
		}
		try {
			// 从内存中写入文件中
			wwb.write();
			// 关闭资源，释放内存
			wwb.close();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}
}
