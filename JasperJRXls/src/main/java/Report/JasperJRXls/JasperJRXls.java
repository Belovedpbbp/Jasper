package Report.JasperJRXls;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import org.codehaus.groovy.tools.shell.commands.ShowCommand;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JRExporterParameter;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRCsvDataSource;
import net.sf.jasperreports.engine.data.JRXlsDataSource;
import net.sf.jasperreports.engine.export.JRXlsExporter;
import net.sf.jasperreports.engine.util.ClassLoaderResource;
import net.sf.jasperreports.view.JasperViewer;

public class JasperJRXls {

	public static JasperPrint printReport() {

		InputStream reportFileName = ClassLoader.class

		.getResourceAsStream("/blank1.jrxml");

		InputStream xlsFileName = ClassLoader.class
				.getResourceAsStream("/data.xls");

		JasperReport sourceFileName;
		JasperPrint jasperPrint = null;

		try {
			sourceFileName = JasperCompileManager.compileReport(reportFileName);
			JRXlsDataSource source = new JRXlsDataSource(xlsFileName);
			source.setColumnNames(new String[] { "name", "ages", "like" },
					new int[] { 0, 1, 2 });

			Map parameters = new HashMap();
			parameters.put("Department", "EssoOil");
			parameters.put("Location", "Chonburi");

			jasperPrint = JasperFillManager.fillReport(sourceFileName,
					parameters, source);

		} catch (JRException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return jasperPrint;

	}

	public static JasperExportManager exportManager() {

		printReport();
		try {
			JasperExportManager.exportReportToHtmlFile(printReport(),
					"D:/Coop/JasperReport/Export/nnoob.html");
		} catch (JRException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// return exportManager();
		return exportManager();

	}

	// public static JasperViewer viewReport() {
	// printReport();
	//
	// JasperViewer.viewReport(printReport(), false);
	// // return viewReport();
	// return viewReport();
	// }

	public static void main(String[] args) {
		exportManager();

		// viewReport();

	}

}
