package a;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFChildAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorksheetDocument;

public class A1 extends HttpServlet {

	/**
	 * Constructor of the object.
	 */
	public A1() {
		super();
	}

	/**
	 * Destruction of the servlet. <br>
	 */
	public void destroy() {
		super.destroy(); // Just puts "destroy" string in log
		// Put your code here
	}

	/**
	 * The doGet method of the servlet. <br>
	 *
	 * This method is called when a form has its tag value method equals to get.
	 * 
	 * @param request the request send by the client to the server
	 * @param response the response send by the server to the client
	 * @throws ServletException if an error occurred
	 * @throws IOException if an error occurred
	 */
	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		response.setContentType("text/html");
		response.setCharacterEncoding("utf-8");
		PrintWriter out = response.getWriter();
		out.println("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\">");
		out.println("<HTML>");
		out.println("  <HEAD><TITLE>A Servlet</TITLE></HEAD>");
		out.println("  <BODY>");
	
		//读取account.xls这个表
		
		//1、把excel文件，变成一个能够被程序处理的数据对象
		File file = new File("/Users/You/Desktop/软件开发/20180622_Servlet&读取Excel/accounts.xls");
		
		//2、在文件对象上插一根数据管道，吧数据抽出来
		InputStream is = new FileInputStream(file);
		
		//3、创建一个工作表对象
		Workbook wb = new HSSFWorkbook(is);
		
		//4、获取工作表的第一个sheet
		Sheet sheet = wb.getSheetAt(0);
		
		//5、拿到数据的开始行序号，和结束行序号
		int firstRowIndex = sheet.getFirstRowNum();
		int lastRowIndex = sheet.getLastRowNum();
		
		//out.println(firstRowIndex + "...." + lastRowIndex);
		
		//6、循环获取每一个行的数据 
		out.println("<table border='1'>");
		for(int i=firstRowIndex ; i<=lastRowIndex; i++)
		{
			//是根据有有效的行号，拿到对应一行的数据
			Row row = sheet.getRow(i);
			
			if(row != null)
			{
				out.println("<tr>");
				//7、读取一行的每一个有效单元格（Cell）的数据
				int firstCellIndex = row.getFirstCellNum();
				int lastCellIndex = row.getLastCellNum();
				
				//8、根据有效的列的范围，读取每一个有效单元格的值
				for(int j=firstCellIndex; j<lastCellIndex; j++)
				{
					//根据列的编号，在一个特定的行中取出特定单元格的值
					//String value = row.getCell(j).toString();
					Cell cell = row.getCell(j);
					out.println("<td>" + cell + "</td>");
					
				}
				
				out.println("</tr>");
			}
			
			out.println("<br>");
			
			
		}
		out.println("</talbe>");
		
		is.close();
		out.println("<a href='a1.jsp' type='button'>添加用户</a>");
		out.println("<br>");
		out.println("<a href='' type='button'>返回json</a>");
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
