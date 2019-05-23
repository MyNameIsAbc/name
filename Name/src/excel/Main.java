package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Main {
	static File logfile;
	static Date a;
	static Calendar c = Calendar.getInstance();
	static int year = c.get(Calendar.YEAR);
	static int month = c.get(Calendar.MONTH);
	static int date = c.get(Calendar.DATE);
	private static ExecutorService cachedThreadPool = Executors.newCachedThreadPool();
	static File outFileParentFile2 = null;
    public static String rubbishPath="E:" + File.separator + "match" + File.separator + "offical"+File.separator+"rubbish";
	public static void main(String[] args) {
		logfile = new File("E:" + File.separator + "match" + File.separator + "time.txt");
		if (!logfile.exists()) {
			try {
				logfile.createNewFile();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		CopyOnWriteArrayList<File> nameFiles = new CopyOnWriteArrayList<File>();
		Utils.getFileList(nameFiles, "E:" + File.separator + "match" + File.separator + "name");
		CopyOnWriteArrayList<File> officalFiles = new CopyOnWriteArrayList<File>();
		Utils.getFileList(officalFiles, "E:" + File.separator + "match" + File.separator + "offical",rubbishPath);

		if (nameFiles.isEmpty()) {
			Utils.print("不存在name表！........退出", logfile);
		} else {
			System.out.println(nameFiles.size());
			initFile();
		
			for (File namefile : nameFiles) {
				List<ModelName> modelNames = new ArrayList<ModelName>();
				modelNames.addAll(startFindname(namefile));
				Utils.print("开始录入name表的数据！", logfile);
				Utils.print("录入name表" + namefile.getName() + "的数据完毕！共有" + modelNames.size() + "个数据", logfile);
				File outFileParentFile3 = new File(outFileParentFile2, namefile.getName().replaceAll(".xls", ""));
				outFileParentFile3.mkdirs();
				for (File file : officalFiles) {
					File outfile = new File(outFileParentFile3, file.getName().replaceAll(".xls", ""));
					System.out.println("outfile:" + outfile.getAbsolutePath());
					System.out.println(outfile.mkdirs());
					List<Result> results = new ArrayList<>();
					List<ModelOffical> modelOfficals = new ArrayList<>();
					modelOfficals.addAll(findOffical(file));
					Utils.print(file.getName() + "有" + modelOfficals.size() + "个数据", logfile);

					for (ModelOffical modelOffical : modelOfficals) {
						if (modelOffical.getName().isEmpty()) {
							continue;
						}
						for (ModelName modelName : modelNames) {
							if (modelName.getName().isEmpty()) {
								continue;
							}
							if (modelName.getName().equalsIgnoreCase(modelOffical.getName())) {
								extractResult(results, modelOffical, modelName);
							}
						}
					}

					if (!results.isEmpty()) {
						cachedThreadPool.execute(new Runnable() {
							@Override
							public void run() {
								outFile(file, outfile, results);
							}
						});

					}

				}
			}
		}
		cachedThreadPool.shutdown();
		if (cachedThreadPool.isTerminated()) {
			Utils.stopprint();
		}
	}

	private static void initFile() {
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH时" + "mm分");
		Date date = new Date();
		String timeString = format.format(date);
		File outFileParentFile = new File("E:" + File.separator + "match" + File.separator + "nameout");
		if (!outFileParentFile.exists()) {
			outFileParentFile.mkdir();
		}
		outFileParentFile2 = new File(outFileParentFile, timeString);
		outFileParentFile2.mkdirs();
	}

	private static void extractResult(List<Result> results, ModelOffical modelOffical, ModelName modelName) {
		Result result = new Result(modelOffical.name, modelName.mobile,
				"未报名,学员  " + modelOffical.name + " 招考岗位 ：" + modelOffical.aginent + modelOffical.findjob + " ,职位代码："
						+ modelOffical.finjobcode + "  考了" + modelOffical.finalscore + "分  " + "  招: "
						+ modelOffical.finjobnumber + "人  " + "排:" + modelOffical.rank + "名  ",
				" ");
		results.add(result);
	}

	// 把name表的数据录进来
	private static List<ModelName> startFindname(File file) {
		InputStream inputStream = null;
		Workbook workbook = null;
		List<ModelName> modelNames = new ArrayList<>();
		try {
			inputStream = new FileInputStream(file);
			workbook = Workbook.getWorkbook(inputStream);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		Sheet sheet = workbook.getSheet(0);
		int name = -1, phone = -1, row = -1, namecount = 0, phonecount = 0;

		for (int j = 0; j < sheet.getRows(); j++) {
			for (int i = 0; i < sheet.getColumns(); i++) {
				String columnname = sheet.getCell(i, j).getContents();
				columnname = columnname.replaceAll("\r|\n", "");
				columnname = columnname.replaceAll(" ", "");
				if (columnname.contains("姓名")) {
					if (namecount == 0) {
						row = j;
						name = i;
					}
					namecount++;
				} else if (columnname.contains("手机号")) {
					if (phonecount == 0) {
						phone = i;
					}
					phonecount++;
				}
				if (name != -1 && phone != -1) {
					break;
				}
			}
			if (name != -1 && phone != -1) {
				break;
			}
		}
		if (name == -1 && phone == -1) {
			name = 0;
			phone = 1;
			row = 0;
		}
		if (row != -1) {
			for (int i = row + 1; i < sheet.getRows(); i++) {
				ModelName modelName = new ModelName();
				if (name != -1) {
					modelName.setName(sheet.getCell(name, i).getContents());
				}
				if (phone != -1) {
					modelName.setMobile(sheet.getCell(phone, i).getContents());
				}

				modelNames.add(modelName);
			}
		}
		System.out.println("namesize:" + modelNames.size());
		System.out.println(modelNames.get(0).toString());
		return modelNames;

	}

	private static List<ModelOffical> findOffical(File file) {
		System.out.println("find offical:" + file.getAbsolutePath());
		List<ModelOffical> modelOfficals = new ArrayList<>();
		Workbook workbook;
		// 获取Excel文件对象
		try {
			InputStream inputStream = new FileInputStream(file);
			workbook = Workbook.getWorkbook(inputStream);
			// 获取文件的指定工作表 默认的第一个
			int number = 0;
			if (workbook.getNumberOfSheets() == 1) {
				number = 0;
				modelOfficals = findOfficalSheet(file, workbook, 0);
				System.out.println("--------------找到官方表" + file.getName() + "在第" + number + "个子表");
			} else {
				String names[] = workbook.getSheetNames();
				System.out.println("--------------names[] " + names.toString());
				for (int i = 0; i < names.length; i++) {
					System.out.println("表单名:" + names[i]);
					number = i;
					modelOfficals.addAll(findOfficalSheet(file, workbook, number));
					System.out.println("--------------找到官方表" + file.getName() + "在第" + number + "个子表");
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return modelOfficals;
	}

	private static List<ModelOffical> findOfficalSheet(File file, Workbook workbook, int number) {
		Sheet sheet = workbook.getSheet(number);
		List<ModelOffical> modelOfficals = new ArrayList<>();
		System.out.println("Colums:" + sheet.getColumns() + "Rows:" + sheet.getRows());
		int agent = -1, findjob = -1, findjobcode = -1, findjobnumber = -1, name = -1, examcode = -1, xingcescore = -1,
				shenglunscore = -1, finalscore = -1, rank = -1;
		int row = -1;
		int finalrow = sheet.getRows();
		if (sheet.getRows()==0|| sheet.getColumns()==0) {
			return modelOfficals;
		}
		for (int j = 0; j < sheet.getRows(); j++) {
			for (int i = 0; i < sheet.getColumns(); i++) {
				String Columname = sheet.getCell(i, j).getContents();
				Columname = Columname.replaceAll("\r|\n", "");
				Columname = Columname.replaceAll(" ", "");
				if (Columname.contains("招录机关")) {
					row = j;
					agent = i;
				} else if (Columname.contains("招录职位")) {
					row = j;
					findjob = i;
				} else if (Columname.contains("职位代码")) {
					row = j;
					findjobcode = i;
				} else if (Columname.contains("招录计划") || Columname.contains("招考人数")
						|| Columname.contains("招录" + "\n" + "计划")) {
					row = j;
					findjobnumber = i;
				} else if (Columname.contains("姓名")) {
					row = j;
					name = i;
				} else if (Columname.contains("准考证号") || Columname.contains("考号")) {
					row = j;
					examcode = i;
				} else if (Columname.contains("行测") || Columname.contains("行政职业能力测验")) {
					row = j;
					xingcescore = i;
				} else if (Columname.contains("申论")) {
					row = j;
					shenglunscore = i;
				} else if (Columname.contains("笔试折算成绩") || Columname.contains("笔试折算分") || Columname.contains("折算分")) {
					row = j;
					finalscore = i;
				} else if (Columname.contains("笔试排名") || Columname.contains("笔试成绩排名") || Columname.contains("排序")
						|| Columname.contains("排名")) {
					row = j;
					rank = i;
				}
				if (findjob != -1 && findjobcode != -1 && name != -1 && finalscore != -1 && rank != -1
						&& findjobnumber != -1) {
					break;
				}
			}
			if (findjob != -1 && findjobcode != -1 && name != -1 && finalscore != -1 && rank != -1
					&& findjobnumber != -1) {
				break;
			}
		}
		System.out.println(findjob+" "+findjobcode+" "+name+" "+finalscore+" "+rank+" "+findjobnumber);
		if (!(findjob != -1 && findjobcode != -1 && name != -1 && finalscore != -1 && rank != -1
				&& findjobnumber != -1)) {
			if (findjob == -1) {
				Utils.print(file.getName() + "Offical表找不到招录职位,请改为招录职位", logfile);
			}
			if (findjobcode == -1) {
				Utils.print(file.getName() + "Offical表找不到职位代码,请改为职位代码", logfile);
			}
			if (name == -1) {
				Utils.print(file.getName() + "Offical表找不到名字,请改为名字", logfile);
			}
			if (finalscore == -1) {
				Utils.print(file.getName() + "Offical表找不到分数,请改为折算分", logfile);
			}
			if (rank == -1) {
				Utils.print(file.getName() + "Offical表找不到排名,请改为排名", logfile);
			}
			if (findjobnumber == -1) {
				Utils.print(file.getName() + "Offical表找不到招录计划,请改为职位代码", logfile);
			}

		} else {
			for (int j = row + 1; j < finalrow; j++) {
				ModelOffical book = new ModelOffical();
				if (agent != -1) {
					book.setAginent(sheet.getCell(agent, j).getContents());
					if (book.getAginent().isEmpty()) {
						book.setAginent(modelOfficals.get(modelOfficals.size() - 1).getAginent());
					}
				}
				if (findjob != -1) {
					book.setFindjob(sheet.getCell(findjob, j).getContents());
					if (book.getFindjob().isEmpty()) {
						book.setFindjob(modelOfficals.get(modelOfficals.size() - 1).getFindjob());
					}
				}
				if (findjobcode != -1) {
					book.setFinjobcode(sheet.getCell(findjobcode, j).getContents());
					if (book.getFinjobcode().isEmpty()) {
						book.setFinjobcode(modelOfficals.get(modelOfficals.size() - 1).getFinjobcode());
					}
				}
				if (findjobnumber != -1) {
					if (sheet.getCell(findjobnumber, j).getContents().length() >= 2) {
						String string = sheet.getCell(findjobnumber, j).getContents();
						string = string.substring(0, string.indexOf("人"));
						book.setFinjobnumber(string);
					} else {
						book.setFinjobnumber(sheet.getCell(findjobnumber, j).getContents());
					}
					if (book.getFinjobnumber().isEmpty()) {
						book.setFinjobnumber(modelOfficals.get(modelOfficals.size() - 1).getFinjobnumber());
					}
				}
				if (name != -1) {
					book.setName(sheet.getCell(name, j).getContents());
				}
				if (examcode != -1) {
					book.setExamcode(sheet.getCell(examcode, j).getContents());

				}
				if (xingcescore != -1) {
					book.setXingcescore(sheet.getCell(xingcescore, j).getContents());

				}
				if (shenglunscore != -1) {
					book.setShenglunscore(sheet.getCell(shenglunscore, j).getContents());

				}
				if (finalscore != -1) {
					book.setFinalscore(sheet.getCell(finalscore, j).getContents());

				}
				if (rank != -1) {
					book.setRank(sheet.getCell(rank, j).getContents());

				}
				modelOfficals.add(book);
			}
		}
		return modelOfficals;

	}

	public synchronized static void exportExcel(String fileName, List<Result> list) {
		WritableWorkbook wwb;
		FileOutputStream fos;
		try {
			fos = new FileOutputStream(fileName);
			wwb = Workbook.createWorkbook(fos);
			WritableSheet ws = wwb.createSheet("结果", 10); // 创建一个工作表
			ws.setColumnView(0, 20);
			ws.setColumnView(1, 20);
			ws.setColumnView(2, 130);
			ws.setColumnView(3, 70);
			// 设置单元格的文字格式
			WritableFont wf = new WritableFont(WritableFont.ARIAL, 11, WritableFont.NO_BOLD, false,
					UnderlineStyle.NO_UNDERLINE, Colour.GREEN);
			WritableCellFormat wcf = new WritableCellFormat(wf);
			wcf.setVerticalAlignment(VerticalAlignment.CENTRE);
			wcf.setAlignment(Alignment.CENTRE);
			ws.setRowView(1, 500);

			WritableFont wf2 = new WritableFont(WritableFont.ARIAL, 15, WritableFont.BOLD, false,
					UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
			WritableCellFormat wcf2 = new WritableCellFormat(wf2);
			wcf2.setVerticalAlignment(VerticalAlignment.CENTRE);
			wcf2.setAlignment(Alignment.CENTRE);
			ws.setRowView(1, 500);

			jxl.write.Label jlabel1 = new jxl.write.Label(0, 0, "姓名");
			jxl.write.Label jlabel2 = new jxl.write.Label(1, 0, "手机号码");
			jxl.write.Label jlabel3 = new jxl.write.Label(2, 0, "描述");
			jxl.write.Label jlabel4 = new jxl.write.Label(3, 0, "解决方案");
			jlabel1.setCellFormat(wcf2);
			jlabel2.setCellFormat(wcf2);
			jlabel3.setCellFormat(wcf2);
			jlabel4.setCellFormat(wcf2);
			ws.addCell(jlabel1);
			ws.addCell(jlabel2);
			ws.addCell(jlabel3);
			ws.addCell(jlabel4);
			// 填充数据的内容

			for (int i = 0; i < list.size(); i++) {
				Result result = list.get(i);
				jxl.write.Label label = new jxl.write.Label(0, i + 1, result.name);
				label.setCellFormat(wcf);
				jxl.write.Label label2 = new jxl.write.Label(1, i + 1, result.phone);
				label2.setCellFormat(wcf);
				jxl.write.Label labe3 = new jxl.write.Label(2, i + 1, result.msg1);
				labe3.setCellFormat(wcf);
				jxl.write.Label label4 = new jxl.write.Label(3, i + 1, result.msg2);
				label4.setCellFormat(wcf);
				ws.addCell(label);
				ws.addCell(label2);
				ws.addCell(labe3);
				ws.addCell(label4);
			}

			wwb.write();
			wwb.close();
			Date bDate = new Date();
			long interval = (bDate.getTime() - a.getTime()) / 1000;
			SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			String txt = "   用时:" + interval + "秒";

			Utils.print("开始时间:" + df.format(a), logfile);
			Utils.print("结束时间:" + df.format(bDate), logfile);
			Utils.print(txt, logfile);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void outFile(File file, File outfile, List<Result> results) {
		File fresultfile1 = new File(outfile, file.getName().replaceAll(".xls", "") + "out.xls");
		if (!fresultfile1.exists()) {
			try {
				fresultfile1.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		Utils.print(file.getName() + "匹配到" + results.size() + "个数据,开始导出", file);
		a = new Date();
		int number = results.size();
		if (number > 0) {
			if (number > 1000) {
				int plus = 0;
				if (number % 3 > 0) {
					plus = (number - number % 3) / 3;
				} else {
					plus = number / 3;
				}
				List<Result> list1 = new ArrayList<>();
				List<Result> list2 = new ArrayList<>();
				List<Result> list3 = new ArrayList<>();

				for (int i = 0; i < plus; i++) {
					list1.add(results.get(i));
				}
				for (int i = plus; i < plus * 2; i++) {
					list2.add(results.get(i));
				}
				for (int i = plus * 2; i < plus * 3; i++) {
					list3.add(results.get(i));
				}

				System.out.println("number:" + number + "plus:" + plus);
				System.out.println("1:" + list1.size());
				System.out.println("2:" + list2.size());
				System.out.println("3:" + list3.size());

				exportExcel(fresultfile1.getAbsolutePath(), list1);

				File fresultfile2 = new File(outfile, file.getName().replaceAll(".xls", "") + "out2.xls");
				if (!fresultfile2.exists()) {
					try {
						fresultfile2.createNewFile();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}

				File fresultfile3 = new File(outfile, file.getName().replaceAll(".xls", "") + "out3.xls");
				if (!fresultfile3.exists()) {
					try {
						fresultfile3.createNewFile();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
				exportExcel(fresultfile2.getAbsolutePath(), list2);
				exportExcel(fresultfile3.getAbsolutePath(), list3);

			} else {
				exportExcel(fresultfile1.getAbsolutePath(), results);
			}
		}
	}
}
