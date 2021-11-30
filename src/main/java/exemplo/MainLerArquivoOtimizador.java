package exemplo;

import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import leitor.LeitorRdhXls;
import leitorexcel.LeitorExcel;
import leitorexcel.LeitorExcel.EXTENSAO;
import utils.ResourceUtil;

public class MainLerArquivoOtimizador {

	public static void main(String[] args) throws Exception {
		System.out.println("INICIO DO PROGRAMA");
		
		// LENDO DOS RESOURCES
		InputStream is = ResourceUtil.readFile("AAOtmizador.xlsm");
		LeitorExcel leitor = new LeitorExcel(is, EXTENSAO.XLSM);

		// LENDO DO ARQUIVO
		//		LeitorExcel leitor = new LeitorExcel("c:\\tmp\\AAOtmizador.xlsm");
		
		leitor.abrir("PLD");
		double valorc2 = leitor.getCelulaDouble("C2");
		System.out.println("CELULA C2=" + valorc2);
		
		leitor.abrir("DadosDePlaca");
		double volumeMaximo = leitor.getCelulaDouble("B2");
		System.out.println("Volume Maximo=" + volumeMaximo);
		
		double polinCota = leitor.getCelulaDouble("B23");
		System.out.println("PolinCota=" + polinCota);
		
		leitor.abrir("CotaVolume");
		System.out.println("--- VOLUMES ---");
		List<Double> volumes = leitor.getColunaDouble("B4:B23");
		volumes.forEach(System.out::println);
	}
}
