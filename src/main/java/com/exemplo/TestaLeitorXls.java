package com.exemplo;


import java.io.IOException;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

public class TestaLeitorXls {

	public static void main(String[] args) throws Exception {
		
		LeitorExcel leitor = new LeitorExcel(Paths.get("c:/temp/RDH_09OUT2021.xlsx"));
		
		String celulaAlterar = "A1";
		double valor = 111.0;
		
		//leitor.abrir(0);
		
		String nomePlanilha = leitor.listarPlanilhas().stream()
				.filter(nome -> nome.contains("Hidrológica"))
				.findFirst().get();

		System.out.println("nome planilha="+ nomePlanilha);
		leitor.abrir(nomePlanilha);
		System.out.println(leitor);
		
//		List<Double> valores = leitor.getValoresColunaDouble(5, 9, 23);
//		valores.forEach(System.out::println);
		
		//List<Double> valores = leitor.getValoresLinhaDouble(6, 13, 9);
		List<Double> valores = leitor.getValoresLinhaDouble("F9:M9");
		valores.forEach(System.out::println);
		
		leitor.setValorCelulaDouble("N9", 12345.0);
		System.out.println("=>" + leitor.getValorCelulaDouble("N9"));
		
		//leitor.setValorCelulaDouble(celulaAlterar, valor);
		//leitor.setValorCelulaBoolean(2,1, false);
		
		//System.out.println("F10=" + leitor.getValorCelulaDouble("F10"));
		//leitor.setValorCelulaDouble("J9", 1239.0);
		//System.out.println(leitor.toString());
		//System.out.println(leitor.getValorCelulaDouble(5, 11));
		
		leitor.gravar(Paths.get("c:\\temp\\anovoArquivo.xlsx"));
		System.out.println("--- FIM ---");
	}

}
