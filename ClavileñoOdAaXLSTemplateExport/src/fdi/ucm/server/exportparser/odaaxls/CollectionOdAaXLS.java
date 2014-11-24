/**
 * 
 */
package fdi.ucm.server.exportparser.odaaxls;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollectionLog;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteLinkElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementURL;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteIterator;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteLinkElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteResourceElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteStructure;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

/**
 * @author Joaquin Gayoso-Cabada
 *Clase qie produce el XLSI
 */
public class CollectionOdAaXLS {


	public static String processCompleteCollection(CompleteCollectionLog cL,
			CompleteCollection salvar, boolean soloEstructura, String pathTemporalFiles) throws IOException {
		
		 /*La ruta donde se creará el archivo*/
        String rutaArchivo = pathTemporalFiles+"/"+System.nanoTime()+".xls";
        /*Se crea el objeto de tipo File con la ruta del archivo*/
        File archivoXLS = new File(rutaArchivo);
        /*Si el archivo existe se elimina*/
        if(archivoXLS.exists()) archivoXLS.delete();
        /*Se crea el archivo*/
        archivoXLS.createNewFile();
        
        /*Se crea el libro de excel usando el objeto de tipo Workbook*/
        Workbook libro = new HSSFWorkbook();
        
        /*Se inicializa el flujo de datos con el archivo xls*/
        FileOutputStream archivo = new FileOutputStream(archivoXLS);
        
        /*Utilizamos la clase Sheet para crear una nueva hoja de trabajo dentro del libro que creamos anteriormente*/
        
        HashMap<Long, Integer> clave=new HashMap<Long, Integer>();	
        
//        Sheet hoja;
        
       
        
        Sheet HojaD = libro.createSheet("Datos");
	   	Sheet HojaM =libro.createSheet("MetaDatos");
	   	Sheet HojaR =libro.createSheet("Recursos");
	   	Sheet HojaF =libro.createSheet("Archivos");
	   	Sheet HojaU =libro.createSheet("URLs");
//	   	Sheet HojaFUR =libro.createSheet("Relacion Recursos");
       
	   	CompleteGrammar Files=findFiles(salvar.getMetamodelGrammar());
	    if (Files!=null)
        {
	    		List<CompleteDocuments> ListaDocumentos=generaDocs(salvar.getEstructuras(),Files);
	    		processFiles(HojaF,Files,clave,cL,ListaDocumentos,soloEstructura);
        		 
        	 


        }
        else
        	cL.getLogLines().add("No posee Files");
	   	
	    CompleteGrammar URLS=findURLs(salvar.getMetamodelGrammar());
	    if (URLS!=null)
        {
	    		List<CompleteDocuments> ListaDocumentos=generaDocs(salvar.getEstructuras(),URLS);
	    		processURls(HojaU,URLS,clave,cL,ListaDocumentos,soloEstructura);
        		 
        	 


        }
        else
        	cL.getLogLines().add("No posee URLs");
	    
	   	
	    CompleteGrammar VirtualObject=findVO(salvar.getMetamodelGrammar());
        if (VirtualObject!=null)
        {
        	 CompleteElementType Datos=findDatos(VirtualObject.getSons());
        	 CompleteElementType MetaDatos=findMetaDatos(VirtualObject.getSons());
        	 CompleteElementType Recursos=findResources(VirtualObject.getSons());
//        	 CompleteElementType RecursosLink=findResourcesLink(VirtualObject.getSons());
        	 
        	 List<CompleteDocuments> ListaDocumentos=generaDocs(salvar.getEstructuras(),VirtualObject);
        	 
        		 processDatos(HojaD,Datos,clave,cL,ListaDocumentos,soloEstructura,VirtualObject.getClavilenoid());
        		 
        	 
        	 if (MetaDatos!=null)
        		 processMetadatos(HojaM,MetaDatos,clave,cL,ListaDocumentos,soloEstructura);

        	 
        	 if (Recursos!=null)
        		 processRecursos(HojaR,Recursos,clave,cL,ListaDocumentos,soloEstructura);
        	 
//        	 if (RecursosLink!=null)
//        		 processRecursosLink(HojaFUR,RecursosLink,clave,cL,ListaDocumentos,soloEstructura);

        }
        else
        	cL.getLogLines().add("No posee un Objeto Virtual");
			
        
        
        
        
//        /*Escribimos en el libro*/
        libro.write(archivo);
        /*Cerramos el flujo de datos*/
        archivo.close();
        /*Y abrimos el archivo con la clase Desktop*/
//        Desktop.getDesktop().open(archivoXLS);
		return rutaArchivo;
//        }
//        else 
//        	{
//        	 libro.write(archivo);
//        	archivo.close();
//        	return "";
//        	}
        

        
        
	}
	
	  private static CompleteElementType findDatos(ArrayList<CompleteStructure> sons) {
		  for (CompleteStructure completeStruct : sons) {
				if (completeStruct instanceof CompleteElementType && StaticFuctionsOdAaXLS.isDatos((CompleteElementType)completeStruct))
					return (CompleteElementType)completeStruct;
			}
			return null;
	}
	  
	  private static CompleteElementType findMetaDatos(ArrayList<CompleteStructure> sons) {
		  for (CompleteStructure completeStruct : sons) {
				if (completeStruct instanceof CompleteElementType && StaticFuctionsOdAaXLS.isMetaDatos((CompleteElementType)completeStruct))
					return (CompleteElementType)completeStruct;
			}
			return null;
	}

	  private static CompleteElementType findResources(ArrayList<CompleteStructure> sons) {
		  
		  for (CompleteStructure completeStruct : sons) {
			  if (completeStruct instanceof CompleteIterator)
				  for (CompleteStructure completeStruct2 : completeStruct.getSons()) {
						if (completeStruct2 instanceof CompleteElementType && StaticFuctionsOdAaXLS.isResources((CompleteElementType)completeStruct2))
							return (CompleteElementType)completeStruct2;
					}
		}
		  
		  
			return null;
	}
	  
//	  private static CompleteElementType findResourcesLink(List<CompleteStructure> sons) {
//		  for (CompleteStructure completeStruct : sons) {
//			  	if (completeStruct instanceof CompleteIterator)
//			  		{
//			  		CompleteElementType TMP=findResourcesLink(completeStruct.getSons());
//			  		if (TMP!=null)
//			  			return TMP;
//			  		}
//			  	else if (completeStruct instanceof CompleteElementType && StaticFuctionsOdAaXLS.isResources((CompleteElementType)completeStruct))
//			  		return (CompleteElementType)completeStruct;
//			}
//			return null;
//	}
	  
	  
	 

	private static CompleteGrammar findVO(List<CompleteGrammar> metamodelGrammar) {
		for (CompleteGrammar completeGrammar : metamodelGrammar) {
			if (StaticFuctionsOdAaXLS.isVirtualObject(completeGrammar))
				return completeGrammar;
		}
		return null;
	}

	private static CompleteGrammar findFiles(List<CompleteGrammar> metamodelGrammar) {
		for (CompleteGrammar completeGrammar : metamodelGrammar) {
			if (StaticFuctionsOdAaXLS.isFiles(completeGrammar))
				return completeGrammar;
		}
		return null;
	}
	
	private static CompleteGrammar findURLs(List<CompleteGrammar> metamodelGrammar) {
		for (CompleteGrammar completeGrammar : metamodelGrammar) {
			if (StaticFuctionsOdAaXLS.isURL(completeGrammar))
				return completeGrammar;
		}
		return null;
	}
	
	private static CompleteElementType findFilesFisico(List<CompleteStructure> metaStructures) {
		for (CompleteStructure completeStructure : metaStructures) {
			if (completeStructure instanceof CompleteElementType&&StaticFuctionsOdAaXLS.isFileFisico((CompleteElementType)completeStructure))
				return (CompleteElementType) completeStructure;
		}
		return null;
	}
	
	private static CompleteElementType findOwner(List<CompleteStructure> metaStructures) {
		for (CompleteStructure completeStructure : metaStructures) {
			if (completeStructure instanceof CompleteElementType&&StaticFuctionsOdAaXLS.isOwner((CompleteElementType)completeStructure))
				return (CompleteElementType) completeStructure;
		}
		return null;
	}
	
	private static CompleteElementType findURI(List<CompleteStructure> metaStructures) {
		for (CompleteStructure completeStructure : metaStructures) {
			if (completeStructure instanceof CompleteElementType&&StaticFuctionsOdAaXLS.isURI((CompleteElementType)completeStructure))
				return (CompleteElementType) completeStructure;
		}
		return null;
	}
	
	
	/**
	 * Para los datos
	 * @param hoja
	 * @param grammar
	 * @param clave
	 * @param cL
	 * @param list
	 * @param soloEstructura
	 * @param gramarId 
	 * @param virtualObject
	 */
	private static void processDatos(Sheet hoja, CompleteElementType grammar,
			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura, Long gramarId) {
		  
	  
		List<CompleteElementType> ListaElementos;
	        
		if (grammar!=null)
	        ListaElementos=generaLista(grammar);
		else ListaElementos=new ArrayList<CompleteElementType>();
	        

	        if (ListaElementos.size()>255)
	        	{
	        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls para structura: " + grammar.getName() +" solo 255 estructuras seran grabadas, divide en gramaticas mas simples");
	        	ListaElementos=ListaElementos.subList(0, 254);
	        	}
	        
	        
	      
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 2; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	if (i==0)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Clavy Document Id ( DO NOT MODIFY THIS COLUMN )";
		            	else 
		            		if (j==1)
		            			Value="Description";
		            		else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=pathFather(TmpEle,grammar);
		            		}
		
		            	
		            	if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            		
		            		
		            	if (j>1)
		            		{
		            		clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
		            		Column++;
		            		}
		            	else
		            	{
		            		hoja.setColumnWidth(j, 12750);
		            	}
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	else if (i==1)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		        		if (j==0)
		            		Value="Clavy Type Id ( DO NOT MODIFY THIS ROW )";
		            	else 
		            		if (j==1)
		            			Value=Long.toString(gramarId);
		            		else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=Long.toString(TmpEle.getClavilenoid());
		            		}
		
		            	
		        		if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	
			}	
	        
	        
	        if (!soloEstructura&&ListaElementos.size()>0)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row fila = hoja.createRow(row);
	            row++;

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null)
						{
						ArrayList<CompleteElement> Lis=ListaClave.get(val);
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaClave.put(val, Lis);
						}
				}
	            
	            
	            
	            /*Cada fila tendrá celdas de datos*/
	            for(int c=0;c<columnsMax+2;c++){
	            	
	            	String Value = "";
	            	if (c==0)
	            		Value=Long.toString(Doc.getClavilenoid());
	            	else if (c==1)
	            		Value=Doc.getDescriptionText();
	            	else
	            		{
	            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
	            		if (temp!=null)
	            		{
	            			if (temp.size()>0){
	            			CompleteElement completeElement=temp.get(0);
	            			Value=getValueFromElement(completeElement);
	            		}

	            		
	            		}
	            		}
	
	            	 
	            	if (Value.length()>=32767)
	            	{
	            		Value="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celda = fila.createCell(c);               	
	                		 celda.setCellValue(Value);
	                    /*Si no es la primera fila establecemos un valor*/
	                	//32.767

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}

	
	/**
	 * Para todo lo que no son los datos
	 * @param hoja
	 * @param grammar
	 * @param clave
	 * @param cL
	 * @param list
	 * @param soloEstructura
	 * @param virtualObject
	 */
	private static void processMetadatos(Sheet hoja, CompleteElementType grammar,
			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  

		

		 List<CompleteElementType> ListaElementos=generaLista(grammar);
	        

	        if (ListaElementos.size()>255)
	        	{
	        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls para gramatica: " + grammar.getName() +" solo 255 estructuras seran grabadas, divide en gramaticas mas simples");
	        	ListaElementos=ListaElementos.subList(0, 254);
	        	}
	
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 2; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	if (i==0)
	        	{
	        		for (int j = 0; j < columnsMax+1; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Clavy Document Id ( DO NOT MODIFY THIS COLUMN )";
		            	else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-1);
		            		Value=pathFather(TmpEle,grammar);
		            		}
		
		            	
		            	if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            		
		            		
		            	if (j>0)
		            		{
		            		clave.put(ListaElementos.get(j-1).getClavilenoid(), Column);
		            		Column++;
		            		}
		            	else
		            	{
		            		hoja.setColumnWidth(j, 12750);
		            	}
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	else if (i==1)
	        	{
	        		for (int j = 0; j < columnsMax+1; j++) {
		        		
		        		String Value = "";
		        		if (j==0)
		            		Value="Clavy Type Id ( DO NOT MODIFY THIS ROW )";
		            	else 
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-1);
		            		Value=Long.toString(TmpEle.getClavilenoid());
		            		}
		
		            	
		        		if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	
			}	
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row fila = hoja.createRow(row);
	            row++;

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null)
						{
						ArrayList<CompleteElement> Lis=ListaClave.get(val);
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaClave.put(val, Lis);
						}
				}
	            
	            
	            
	            /*Cada fila tendrá celdas de datos*/
	            for(int c=0;c<columnsMax+1;c++){
	            	
	            	String Value = "";
	            	if (c==0)
	            		Value=Long.toString(Doc.getClavilenoid());
	            	else
	            		{
	            		{
		            		ArrayList<CompleteElement> temp = ListaClave.get(c-1);
		            		if (temp!=null)
		            		{
		            			if (temp.size()>0){
		            			CompleteElement completeElement=temp.get(0);
		            			Value=getValueFromElement(completeElement);
		            			
		            		}

		            		
		            		}
		            		}
	            		}
	
	            	 
	            	if (Value.length()>=32767)
	            	{
	            		Value="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celda = fila.createCell(c);               	
	                		 celda.setCellValue(Value);
	                    /*Si no es la primera fila establecemos un valor*/
	                	//32.767

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}
	
	

	/**
	 * Para todo lo que no son los datos
	 * @param hoja
	 * @param grammar
	 * @param clave
	 * @param cL
	 * @param list
	 * @param soloEstructura
	 * @param virtualObject
	 */
	private static void processRecursos(Sheet hoja, CompleteElementType grammar,
			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  

		

		 List<CompleteElementType> ListaElementos=generaLista(grammar);
	        

	        if (ListaElementos.size()>255)
	        	{
	        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls para gramatica: " + grammar.getName() +" solo 255 estructuras seran grabadas, divide en gramaticas mas simples");
	        	ListaElementos=ListaElementos.subList(0, 254);
	        	}
	
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 2; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	if (i==0)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Clavy Document Id Owner( DO NOT MODIFY THIS COLUMN )";
		            	else if (j==1)
		            		Value="Clavy Document/File/Url Resource Id ( DO NOT MODIFY THIS COLUMN )";
		            	else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=pathFather(TmpEle,grammar);
		            		}
		
		            	
		            	if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            		
		            		
		            	if (j>1)
		            		{
		            		clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
		            		Column++;
		            		}
		            	else
		            	{
		            		hoja.setColumnWidth(j, 12750);
		            	}
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	else if (i==1)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		        		if (j==0)
		            		Value="Clavy Type Id ( DO NOT MODIFY THIS ROW )";
		        		else if (j==1)
		            		Value=Long.toString(grammar.getClavilenoid());
		            	else 
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=Long.toString(TmpEle.getClavilenoid());
		            		}
		
		            	
		        		if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	
			}	
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            
            
	            HashMap<Integer, ArrayList<CompleteElement>> ListaAmbito=new HashMap<Integer, ArrayList<CompleteElement>>();
	            HashMap<Integer,Long> ListaRelacionesporAmbito=new HashMap<Integer,Long>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null&&elem.getAmbitos().size()>0)
						{
						ArrayList<CompleteElement> Lis=ListaAmbito.get(elem.getAmbitos().get(0));
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaAmbito.put(elem.getAmbitos().get(0), Lis);
						}
					
					if (elem instanceof CompleteLinkElement)
					{
						if (((CompleteLinkElement) elem).getValue() != null&&elem.getAmbitos().size()>0)
						{
						ListaRelacionesporAmbito.put(elem.getAmbitos().get(0), ((CompleteLinkElement) elem).getValue().getClavilenoid());	
						}
					}
				}
	            
	            for (Entry<Integer, Long> listarelaciones : ListaRelacionesporAmbito.entrySet()) {
	            	ArrayList<CompleteElement> Elems=ListaAmbito.get(listarelaciones.getKey());
	            	if (Elems==null)
	            		Elems=new ArrayList<CompleteElement>();
	            	
	            	HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
		            for (CompleteElement elem : Elems) {
						Integer val=clave.get(elem.getHastype().getClavilenoid());
						if (val!=null)
							{
							ArrayList<CompleteElement> Lis=ListaClave.get(val);
							if (Lis==null)
								{
								Lis=new ArrayList<CompleteElement>();
								}
							Lis.add(elem);
							ListaClave.put(val, Lis);
							}

					}
		            
		            Row fila = hoja.createRow(row);
		            row++;
		            
		            if (listarelaciones.getValue()!=null)
		            {
		            /*Cada fila tendrá celdas de datos*/
		            for(int c=0;c<columnsMax+2;c++){
		            	
		            	String Value = "";
		            	if (c==0)
		            		Value=Long.toString(Doc.getClavilenoid());
		            	else if (c==1)
		            		Value=Long.toString(listarelaciones.getValue());
		            	else
		            		{
		            		{
			            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
			            		if (temp!=null)
			            		{
			            			if (temp.size()>0){
			            			CompleteElement completeElement=temp.get(0);
			            			Value=getValueFromElement(completeElement);
			            		}

			            		
			            		}
			            		}
		            		}
		
		            	 
		            	if (Value.length()>=32767)
		            	{
		            		Value="";
		            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		                /*Creamos la celda a partir de la fila actual*/
		                Cell celda = fila.createCell(c);               	
		                		 celda.setCellValue(Value);
		                    /*Si no es la primera fila establecemos un valor*/
		                	//32.767

		                
		            	}
				}
	            
	            }
	            
	            
	           
	            
	            
	            
	           

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}
	
	

//	/**
//	 * Para todo lo que no son los datos
//	 * @param hoja
//	 * @param grammar
//	 * @param clave
//	 * @param cL
//	 * @param list
//	 * @param soloEstructura
//	 * @param virtualObject
//	 */
//	private static void processRecursosLink(Sheet hoja, CompleteElementType grammar,
//			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
//		  
//
//
//	
//		
//	        if (ListaDocumentos.size()+2>65536)
//	    	{
//	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
//	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
//	    	}
//
//	        	
//	        int row=0;     	
//	        Row filaPrimera = hoja.createRow(row);
//        	row++;
//        	Cell celda0 = filaPrimera.createCell(0);
//        	celda0.setCellValue("Identificador del Documento");
//        	Cell celda1 = filaPrimera.createCell(1);
//        	celda1.setCellValue("Identificadorede de las URLs/Archivos/Documentos asociados al Documento ----------->");
//        	
//        	
//    		hoja.setColumnWidth(0, 12750);
//
//    	 
//        	
// 
//	        if (!soloEstructura)
//	        {
//	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
//	        for(int f=0;f<ListaDocumentos.size();f++){
//	            /*La clase Row nos permitirá crear las filas*/
//	            Row fila = hoja.createRow(row);
//	            row++;
//
//	            CompleteDocuments Doc=ListaDocumentos.get(f);
//	            
//	            ArrayList<CompleteElement> Elementos=new ArrayList<CompleteElement>();
//	            for (CompleteElement elem : Doc.getDescription()) {
//	            	if (elem.getHastype().getClavilenoid().equals(grammar.getClavilenoid()))
//	            		Elementos.add(elem);
//				}
//	            
//	            
//	            
//	            
//	            /*Cada fila tendrá celdas de datos*/
//	            for(int c=0;c<=Elementos.size();c++){
//	            	
//	            	String Value = "";
//	            	if (c==0)
//	            		Value=Long.toString(Doc.getClavilenoid());
//	            	else
//	            		{
//	            		
//		            		CompleteElement completeElement = Elementos.get(c-1);
//		            		if (completeElement!=null)
//		            		{
//		            			Value=getValueFromElement(completeElement);
//		            		}
//
//		            		
//		            		}
//		            		
//	            		
//	            	
//	            	
//	            	if (Value.length()>=32767)
//	            	{
//	            		Value="";
//	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
//	            		Value.substring(0, 32766);
//	            	}
//	                /*Creamos la celda a partir de la fila actual*/
//	                Cell celda = fila.createCell(c);               	
//	                		 celda.setCellValue(Value);
//	                    /*Si no es la primera fila establecemos un valor*/
//	                	//32.767
//
//	                
//	            	}
//
//	            		
//	            		
//	            }
//	        
//	        }
//	        
//	       
//		
//	}
	
	
	/**
	 * Para todo los files
	 * @param hoja
	 * @param grammar
	 * @param clave
	 * @param cL
	 * @param list
	 * @param soloEstructura
	 * @param virtualObject
	 */
	private static void processFiles(Sheet hoja, CompleteGrammar grammar,
			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  
	  
	        
		  
			List<CompleteElementType> ListaElementos=new ArrayList<CompleteElementType>();
			
			CompleteElementType Filefisico=findFilesFisico(grammar.getSons());
			CompleteElementType FileOwner=findOwner(grammar.getSons());

			Filefisico.setName("Path");
			FileOwner.setName("Identificador del dueño Objeto Virtual");
			
			ListaElementos.add(Filefisico);
			ListaElementos.add(FileOwner);
	      
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 2; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	if (i==0)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Clavy Document Id ( DO NOT MODIFY THIS COLUMN )";
		            	else if (j==1)
		            		Value="Description";
		            	else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=TmpEle.getName();
		            		}
		
		            	
		            	if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            		
		            		
		            	if (j>1)
		            		{
		            		clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
		            		Column++;
		            		}
		            	else
		            	{
		            		hoja.setColumnWidth(j, 12750);
		            	}
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	else if (i==1)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		        		if (j==0)
		            		Value="Clavy Type Id ( DO NOT MODIFY THIS ROW )";
		        		else if (j==1)
		            		Value=Long.toString(grammar.getClavilenoid());
		            	else 
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=Long.toString(TmpEle.getClavilenoid());
		            		}
		
		            	
		        		if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	
			}	
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row fila = hoja.createRow(row);
	            row++;

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null)
						{
						ArrayList<CompleteElement> Lis=ListaClave.get(val);
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaClave.put(val, Lis);
						}
				}
	            
	            
	            
	            /*Cada fila tendrá celdas de datos*/
	            for(int c=0;c<columnsMax+2;c++){
	            	
	            	String Value = "";
	            	if (c==0)
	            		Value=Long.toString(Doc.getClavilenoid());
	            	else if (c==2)
	            		Value=Doc.getDescriptionText();
	            	else
	            		{
	            		{
		            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
		            		if (temp!=null)
		            		{
		            		if (temp.size()>0){
		            			CompleteElement completeElement=temp.get(0);
		            			Value=getValueFromElement(completeElement);
		            		}

		            		
		            		}
		            		}
	            		}
	
	            	 
	            	if (Value.length()>=32767)
	            	{
	            		Value="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celda = fila.createCell(c);               	
	                		 celda.setCellValue(Value);
	                    /*Si no es la primera fila establecemos un valor*/
	                	//32.767

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}
	
	/**
	 * Para todo los files
	 * @param hoja
	 * @param grammar
	 * @param clave
	 * @param cL
	 * @param list
	 * @param soloEstructura
	 * @param virtualObject
	 */
	private static void processURls(Sheet hoja, CompleteGrammar grammar,
			HashMap<Long, Integer> clave, CompleteCollectionLog cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  
	  
	        
		  
			List<CompleteElementType> ListaElementos=new ArrayList<CompleteElementType>();
			
			CompleteElementType URi=findURI(grammar.getSons());

			URi.setName("Path");
			
			ListaElementos.add(URi);
	      
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 2; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	if (i==0)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Clavy Document Id ( DO NOT MODIFY THIS COLUMN )";
		            	else if (j==1)
		            		Value="Description";
		            	else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=TmpEle.getName();
		            		}
		
		            	
		            	if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            		
		            		
		            	if (j>1)
		            		{
		            		clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
		            		Column++;
		            		}
		            	else
		            	{
		            		hoja.setColumnWidth(j, 12750);
		            	}
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	else if (i==1)
	        	{
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		        		if (j==0)
		            		Value="Clavy Type Id ( DO NOT MODIFY THIS ROW )";
		        		else if (j==1)
		            		Value=Long.toString(grammar.getClavilenoid());
		            	else 
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		Value=Long.toString(TmpEle.getClavilenoid());
		            		}
		
		            	
		        		if (Value.length()>=32767)
		            	{
		            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
		            		Value.substring(0, 32766);
		            	}
		            		Cell celda = fila.createCell(j);
		            	
		            	celda.setCellValue(Value);
		            
		           }
	        	}
	        	
			}	
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row fila = hoja.createRow(row);
	            row++;

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null)
						{
						ArrayList<CompleteElement> Lis=ListaClave.get(val);
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaClave.put(val, Lis);
						}
				}
	            
	            
	            
	            /*Cada fila tendrá celdas de datos*/
	            for(int c=0;c<columnsMax+2;c++){
	            	
	            	String Value = "";
	            	if (c==0)
	            		Value=Long.toString(Doc.getClavilenoid());
	            	else if (c==1)
	            		Value=Doc.getDescriptionText();
	            	else
	            		{
	            		{
		            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
		            		if (temp!=null)
		            		{
		            			if (temp.size()>0){
		            			CompleteElement completeElement=temp.get(0);
		            			Value=getValueFromElement(completeElement);
		            		}

		            		
		            		}
		            		}
	            		}
	
	            	 
	            	if (Value.length()>=32767)
	            	{
	            		Value="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celda = fila.createCell(c);               	
	                		 celda.setCellValue(Value);
	                    /*Si no es la primera fila establecemos un valor*/
	                	//32.767

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}
	
	private static ArrayList<CompleteDocuments> generaDocs(
			List<CompleteDocuments> list, CompleteGrammar grammar) {
		ArrayList<CompleteDocuments> ListaDoc=new ArrayList<CompleteDocuments>();
		for (CompleteDocuments completeDocuments : list) {
			if (completeDocuments.getDocument()==grammar)
				ListaDoc.add(completeDocuments);
		}
		return ListaDoc;
	}

//	private static ArrayList<CompleteElementType> generaLista(
//			List<CompleteGrammar> metamodelGrammar) {
//		  ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
//		  for (CompleteGrammar completegramar : metamodelGrammar) {
//			ListaElementos.addAll(generaLista(completegramar));
//		}
//		return ListaElementos;
//	}

	private static ArrayList<CompleteElementType> generaLista(
			CompleteElementType completegramar) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteStructure completeelem : completegramar.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}
	
//	private static ArrayList<CompleteElementType> generaLista(
//			CompleteGrammar completegramar) {
//		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
//		 for (CompleteStructure completeelem : completegramar.getSons()) {
//			 	if (completeelem instanceof CompleteElementType)
//			 		{
//			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
//			 			ListaElementos.add((CompleteElementType)completeelem);
//			 		}
//				ListaElementos.addAll(generaLista(completeelem));
//			}
//		 return ListaElementos;
//	}

	private static Collection<? extends CompleteElementType> generaLista(
			CompleteStructure completeelementPadre) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteStructure completeelem : completeelementPadre.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	public static void main(String[] args) throws Exception{
		
		int id=0;
		
		
		
		  CompleteCollection CC=new CompleteCollection("Lou Arreglate", "Arreglate ya!");
		  for (int i = 0; i < 5; i++) {
			  CompleteGrammar G1 = new CompleteGrammar(new Long(id),"Grammar"+i, i+"", CC);
			  
			  ArrayList<CompleteDocuments> CD=new ArrayList<CompleteDocuments>();
			  int docsN=(new Random()).nextInt(5);
			  docsN=docsN+5;
			for (int j = 0; j < docsN; j++) {
				CompleteDocuments CDDD=new CompleteDocuments(new Long(id), CC, G1, "", "");
				CC.getEstructuras().add(CDDD);
				 id++;
				CD.add(CDDD);
			}
			  
			  id++;
			  for (int j = 0; j < 5; j++) {
				  CompleteElementType CX = new CompleteElementType(new Long(id),"Structure "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteTextElementType CX = new CompleteTextElementType(new Long(id),"Texto "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						CompleteTextElement CTE=new CompleteTextElement(new Long(id), CX, "Texto "+(i*10+j));
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
				
				
				
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteLinkElementType CX = new CompleteLinkElementType(new Long(id),"Link "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						CompleteLinkElement CTE=new CompleteLinkElement(new Long(id), CX, CD.get((new Random()).nextInt(CD.size())));
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteResourceElementType CX = new CompleteResourceElementType(new Long(id),"Resource "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						
						boolean URL=(new Random()).nextBoolean();
						CompleteResourceElement CTE;
						if (URL)
							CTE=new CompleteResourceElementURL(new Long(id), CX, "URL "+(i*10+j));
						else 
							{
							CompleteFile FF = new CompleteFile(new Long(id), "Path File "+(i*10+j), CC);
							CC.getSectionValues().add(FF);
							id++;
							CTE=new CompleteResourceElementFile(new Long(id), CX, FF);
							}
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
				
			}
			  CC.getMetamodelGrammar().add(G1);
		}
		 
		  
		  
		  processCompleteCollection(new CompleteCollectionLog(), CC, false, System.getProperty("user.home"));
		  
	    }


	private static String getValueFromElement(CompleteElement completeElement) {
		try {
			if (completeElement instanceof CompleteTextElement)
    			return (((CompleteTextElement)completeElement).getValue());
			else if (completeElement instanceof CompleteLinkElement)
				return Long.toString((((CompleteLinkElement)completeElement).getValue().getClavilenoid()));
			else if (completeElement instanceof CompleteResourceElementURL)
				return (((CompleteResourceElementURL)completeElement).getValue());
			else if (completeElement instanceof CompleteResourceElementFile)
				return (((CompleteResourceElementFile)completeElement).getValue().getPath());
		} catch (Exception e) {
			return "";
		}
		return "";
	}
	
	/**
	 *  Retorna el Texto que representa al path.
	 * @param grammar 
	 *  @return Texto cadena para el elemento
	 */
	public static String pathFather(CompleteStructure entrada, CompleteElementType grammar)
	{
		String DataShow;
		if (entrada instanceof CompleteElementType)
			DataShow= ((CompleteElementType) entrada).getName();
		else DataShow= "*";
		
		if (entrada.getFather()!=null && entrada.getFather()!=grammar)
			return pathFather(entrada.getFather(),grammar)+"/"+DataShow;
		else if (entrada.getFather()!=null && entrada.getFather()==grammar)
			return DataShow; 
		else return entrada.getCollectionFather().getNombre()+"/"+DataShow;
	}
}
