package a;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.util.ArrayList;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Servlet implementation class A2
 */
@WebServlet("/A2")
public class A2 extends HttpServlet {
	private static final long serialVersionUID = 1L;

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub

	
		response.setContentType("text/html");
		response.setCharacterEncoding("utf-8");
		PrintWriter out = response.getWriter();
		out.println("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\">");
		out.println("<HTML>");
		out.println("  <HEAD><TITLE>A Servlet</TITLE></HEAD>");
		out.println("  <BODY>");
		out.println(" 添加成功 ");
		out.println("<br>");
		out.println(" <a type='button' href='a1.jsp'>返回添加</a> ");
		out.println("<br>");
		out.println(" <a type='button'href='A'>显示</a>  ");
		File file = new File("/Users/You/Desktop/软件开发/20180622_Servlet&读取Excel/accounts.xls");
		InputStream is = new FileInputStream(file);
		
		ArrayList al = new ArrayList();
		
		String a=new String (request.getParameter("a").getBytes("iso-8859-1"),"utf-8");
		String b=new String (request.getParameter("b").getBytes("iso-8859-1"),"utf-8");
		String c=new String (request.getParameter("c").getBytes("iso-8859-1"),"utf-8");
		String d=new String (request.getParameter("d").getBytes("iso-8859-1"),"utf-8");
		String e=new String (request.getParameter("e").getBytes("iso-8859-1"),"utf-8");
		
	
			ArrayList<Test> testList = new ArrayList<Test>();
			Test test = null;
			
			test = new Test();
			test.a = a;
			test.b = b;
			test.c = c;
			test.d = d;
			test.e = e;
			
			testList.add(test);
			
		Workbook wb = new HSSFWorkbook(is);
		
		Sheet sheet = wb.getSheetAt(0);
		
		//找到写入数据的行号
		int rowNum = sheet.getLastRowNum() + 1;
		
		
		//创建新行
		Row row = sheet.createRow(rowNum);
		
		//拿到第一个列的位置
			//我们先拿到第一个有效行，通过这一行来检测出第一个有效列的位置
		Row firstRow = sheet.getRow(sheet.getFirstRowNum());
		int firstCellIndex = firstRow.getFirstCellNum();
		
		Cell cell = null;
		
		cell = row.createCell(firstCellIndex + 0);
		cell.setCellValue(testList.get(0).a);
		cell = row.createCell(firstCellIndex + 1);
		cell.setCellValue(testList.get(0).b);
		
		cell = row.createCell(firstCellIndex + 2);
		cell.setCellValue(testList.get(0).c);
		
		cell = row.createCell(firstCellIndex + 3);
		cell.setCellValue(testList.get(0).d);
		
		cell = row.createCell(firstCellIndex + 4);
		cell.setCellValue(testList.get(0).e);
		
		//创建一个用来写文件的输出流
		OutputStream os = new FileOutputStream(file);
		
		//将变化以后的数据写入文件
		wb.write(os);
		
		
		is.close();
		os.close();
		out.println("  </BODY>");
		out.println("</HTML>");
		out.flush();
		out.close();
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
