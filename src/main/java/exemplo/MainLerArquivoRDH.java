package exemplo;

import java.io.InputStream;

import leitorRDH.LeitorRdhXls;
import leitorexcel.LeitorExcel.EXTENSAO;
import utils.ResourceUtil;

public class MainLerArquivoRDH {

	public static void main(String[] args) throws Exception {
		System.out.println("INICIO DO PROGRAMA");
		
		// EXEMPLO1: LENDO DO ARQUIVO DENTRO DO JAR
		System.out.println("LENDO ARQUIVO RDH_09OUT2021.xlsx DO RESOUCES");
		InputStream is = ResourceUtil.readFile("RDH_09OUT2021.xlsx");   
		LeitorRdhXls rdhExcel = new LeitorRdhXls(is, EXTENSAO.XLSX);
		
		// EXEMPLO2: LENDO DO ARQUIVO EXCEL EXTERNO
//		System.out.println("LENDO DO ARQUIVO RDH_09OUT2021.xlsx");
//		Path pathRdh = Paths.get("c:/temp/RDH_09OUT2021.xlsx");
//		LeitorRdhXls rdhExcel = new LeitorRdhXls(pathRdh);
		
		System.out.format("Usina=%-12s    ValorMesAnteior =%5s\n", 
				rdhExcel.getNomeUsina(2), rdhExcel.getVazaoMesAnterior(2)); 
		
		System.out.format("Usina=%-12s    ValorMesAtual   =%5s\n", 
				rdhExcel.getNomeUsina(2), rdhExcel.getVazaoMesAtual(2)); 
		
		System.out.format("Usina=%-12s    ValorMesAnterior=%5s\n", 
				 rdhExcel.getNomeUsina(211), rdhExcel.getVazaoMesAnterior(211));
		
		System.out.format("Usina=%-12s    ValorMesAtual   =%5s\n", 
				 rdhExcel.getNomeUsina(211), rdhExcel.getVazaoMesAtual(211));
	}
}
