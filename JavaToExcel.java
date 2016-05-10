import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class GenerateExcel {

	/**
	 * @param args
	 */
	public static final String OFFICE_EXCEL_2003_POSTFIX = "xls";
	public static final String OFFICE_EXCEL_2010_POSTFIX = "xlsx";
	public static final String NOT_EXCEL_FILE = " : Not the Excel file!";
	public static final String PROCESSING = "Processing...";
	public static String outputFile = "D:\\test.xls";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			String headers[]={"报告头信息","基本信息","信贷信息概要","信贷信息明细","非信贷交易概要信息","非信贷交易信息明细","公共信息概要","公共信息明细","本人声明","征信中心说明","异议标注(非信贷相关)","查询记录","报告说明信息","编制说明信息"};
			String follows[][]={
					{"报告生成信息","查询请求信息","信息主体标识信息","信息主体防欺诈警示信息","信息主体异议标注信息"},
					{"身份信息","配偶信息","居住信息","职业信息","对外出资记录","企业任职记录","本人声明","异议标注信息"},
					{"信贷提示信息","个人信用报告“数字解读”","逾期及违约信息概要","信贷信息汇总"},
					{"被追偿信息","贷款信息","贷记卡信息","准贷记卡信息","其他账户","关联还款责任信息"},
					{"非信贷交易概要信息"},
					{"电信缴费信息","水费缴费信息","电费缴费信息","煤气缴费信息"},
					{"公共信息概要"},
					{"欠税记录","法院民事判决信息","法院强制执行信息","行政处罚信息","住房公积金缴存信息","养老保险缴存记录","养老保险金发放记录","低保救助信息","执业资格信息","行政奖励信息","车辆交易信息和抵押信息"},
					{"本人声明"},
					{"征信中心说明"},
					{"异议标注信息"},
					{"查询记录汇总信息","政府版社会版查询记录信息明细","不同原因查询信息明细","机构查询信息明细","其他查询记录"},
					{"报告说明信息"},
					{"编制说明信息"}
			};
			/*创建新的Excel 工作簿*/
//			HSSFWorkbook workbook = new HSSFWorkbook();
//			/*在Excel工作簿中建一工作表，其名为缺省值。如要新建一名为"测试情景-新"的工作表，其语句为*/
//			HSSFSheet sheet = workbook.createSheet("测试情景-新");
//			/*在索引0的位置创建行（最顶端的行）*/
//			HSSFRow row = sheet.createRow(0);
//			/*在索引0的位置创建单元格（左上端）*/
//			HSSFCell cell = row.createCell(0);
//			/*定义单元格为字符串类型*/
//			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
//			/*在单元格中输入一些内容*/
//			cell.setCellValue("增加值");
			
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("测试情景-新");
			HSSFRow row;
			HSSFCell cell;
			String cell_string;
			int row_num = 0;
			int colum_num = 0;
			for (int i = 0; i < headers.length; i++) {
				for (int j = 0; j < follows[i].length; j++) {
					row = sheet.createRow(row_num);
					//生成情景编号
					cell=row.createCell(colum_num);
					cell_string=String.format("%s-%s", headers[i],follows[i][j]);
					cell.setCellValue(cell_string);
					//生成情境描述
					colum_num++;
					cell=row.createCell(colum_num);
					cell_string=String.format("1、%s中%s组件中各个组件属性显示是否正确(加工规则)\n2、%s中的%s组件在各个版本报告中的显示规则测试", headers[i],follows[i][j], headers[i],follows[i][j]);
					cell.setCellValue(cell_string);
					row_num++;
					colum_num=0;
				}
			}

			// 新建一输出文件流
			FileOutputStream fOut = new FileOutputStream(outputFile);
			// 把相应的Excel 工作簿存盘
			workbook.write(fOut);
			fOut.flush();
			// 操作结束，关闭文件
			fOut.close();
			System.out.println("文件生成...");
		} catch (Exception e) {
			System.out.println("已运行 xlCreate() : " + e);
		}
	}

}
