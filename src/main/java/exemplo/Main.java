package exemplo;

import java.io.InputStream;

import leitorexcel.LeitorExcel;
import leitorexcel.LeitorExcel.EXTENSAO;
import utils.ResourceUtil;

public class Main {

	public static void main(String[] args) throws Exception {
		System.out.println("LENDO ARQUIVO...");
		InputStream is = ResourceUtil.readFile("RDH_09OUT2021.xlsx");
		LeitorExcel leitor = new LeitorExcel(is, EXTENSAO.XLSX);
		leitor.abrir(0);
		
		// Le um valor usando coluna e linha
		System.out.println("Lendo a célula coluna=6 linha=56");
		System.out.println("   valor= " + leitor.getValorCelulaDouble(6,56));
		
		// Le um valor usando letra e numero
		System.out.println("Lendo a célula F56");
		System.out.println("   valor= " + leitor.getValorCelulaDouble("F56"));
		
		// Le uma coluna de valores
		System.out.println("Lendo uma coluna F48:F58");
		leitor.getValoresColunaDouble("F48:F58").forEach(s -> System.out.println("    " + s));
		
		// Le uma coluna de valores
		System.out.println("Lendo a coluna 6:48-6:58");
		leitor.getValoresColunaDouble(6,48,58).forEach(s -> System.out.println("    " + s));
		
		// Le uma linha de valores
		System.out.println("Lendo a linha F48:M48");
		leitor.getValoresLinhaDouble("F48:M48").forEach(s -> System.out.print("    " + s));
		System.out.println();
		
		// Le uma linha de valores		
		System.out.println("Lendo a linha 6:48-13:48");
		leitor.getValoresLinhaDouble(6,13,48).forEach(s -> System.out.print("    " + s));
		System.out.println();
		
		// Le uma matriz de valores
		{
			System.out.println("Lendo matriz de valores 6:48-7:50");
			Double[][] valores = leitor.getValoresRangeDouble(6, 48, 7, 50);
			for(Double[] linha : valores) {
				for(Double celula : linha) {
					System.out.format("%10s", celula);
				}
				System.out.println();
			}
		}
		
		// Le uma matriz de valores
		{ 
			System.out.println("Lendo matriz de valores F48:G51");
			Double[][] valores = leitor.getValoresRangeDouble("F48:H50");
			for(Double[] linha : valores) {
				for(Double celula : linha) {
					System.out.format("%10s", celula);
				}
				System.out.println();
			}
		}
		System.out.println();
		System.out.println("FIM.");
	}
}
