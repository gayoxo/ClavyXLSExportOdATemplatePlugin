/**
 * 
 */
package fdi.ucm.server.exportparser.odaaxls;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
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
import fdi.ucm.server.modelComplete.collection.CompleteLogAndUpdates;
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


	

	public static String processCompleteCollection(CompleteLogAndUpdates cL,
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
        
//        HashMap<Long, Integer> clave=new HashMap<Long, Integer>();	
        
//        Sheet hoja;
        
       
        
        Sheet HojaD = libro.createSheet(NameConstantsOdAaXLS.DATOS);
	   	Sheet HojaM =libro.createSheet(NameConstantsOdAaXLS.META_DATOS);
	   	Sheet HojaR =libro.createSheet(NameConstantsOdAaXLS.RECURSOS2);
	   	Sheet HojaF =libro.createSheet(NameConstantsOdAaXLS.ARCHIVOS);
	   	Sheet HojaU =libro.createSheet(NameConstantsOdAaXLS.UR_LS);
       
	   	
	   	//Files
	   	List<CompleteDocuments> ListaDocumentosF=new ArrayList<CompleteDocuments>();
	   	CompleteGrammar Files=findFiles(salvar.getMetamodelGrammar());
	   	if (Files==null)
	   	{
	   		cL.getLogLines().add("No posee Files");
	   		Files=new CompleteGrammar("generada", "Generada", salvar);
	   		
	   	}else
	   		ListaDocumentosF=generaDocs(salvar.getEstructuras(),Files);
	   	
	    
	    processFiles(HojaF,Files,
//	    		clave,
	    		cL,ListaDocumentosF,soloEstructura);
        		 
        	 


      
	    List<CompleteDocuments> ListaDocumentosU=new ArrayList<CompleteDocuments>();	
	   	
	    CompleteGrammar URLS=findURLs(salvar.getMetamodelGrammar());
	    if (URLS==null)
        {
	    	cL.getLogLines().add("No posee URLs");
	    	URLS=new CompleteGrammar("generada", "Generada", salvar);
        }
	    else
	    	ListaDocumentosU=generaDocs(salvar.getEstructuras(),URLS);
	    
	    processURls(HojaU,URLS,
//	    		clave,
	    		cL,ListaDocumentosU,soloEstructura);
        		 
        	 


        

	    
	   	
	    CompleteGrammar VirtualObject=findVO(salvar.getMetamodelGrammar());
	    List<CompleteDocuments> ListaDocumentosOV=new ArrayList<CompleteDocuments>();
        if (VirtualObject==null)
        {
        	cL.getLogLines().add("No posee un Objeto Virtual");
        	VirtualObject=new CompleteGrammar("generada", "Generada", salvar);
        }
        else 
        	ListaDocumentosOV=generaDocs(salvar.getEstructuras(),VirtualObject);
        
        ListaDocumentosOV=ordenaDocS(ListaDocumentosOV);
        
        	 CompleteElementType Datos=findDatos(VirtualObject.getSons());
        	 CompleteElementType MetaDatos=findMetaDatos(VirtualObject.getSons());
        	 CompleteElementType Recursos=findResources(VirtualObject.getSons());
//        	 CompleteElementType RecursosLink=findResourcesLink(VirtualObject.getSons());
        	 
        	 if (Datos==null)
 			{
 			cL.getLogLines().add("No posee Datos");
 			Datos=new CompleteElementType();
 			Datos.setClavilenoid(-1l);
 			}
        	 if (MetaDatos==null)
 			{
 			cL.getLogLines().add("No posee Metadatos");
 			MetaDatos=new CompleteElementType();
 			MetaDatos.setClavilenoid(-2l);
 			}
        	 if (Recursos==null)
 			{
 			cL.getLogLines().add("No posee Recursos");
 			Recursos=new CompleteElementType();
 			Recursos.setClavilenoid(-3l);
 			}
        	 
        		 processDatos(HojaD,Datos
//        				 ,clave
        				 ,cL,ListaDocumentosOV,soloEstructura,VirtualObject.getClavilenoid());
        		 processMetadatos(HojaM,MetaDatos
//        				 ,clave
        				 ,cL,ListaDocumentosOV,soloEstructura);
        		 processRecursos(HojaR,Recursos
//        				 ,clave
        				 ,cL,ListaDocumentosOV,soloEstructura);
        	


        	
			
        
        
        
        
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
	
	  private static List<CompleteDocuments> ordenaDocS(
			List<CompleteDocuments> listaDocumentosOV) {
		  quicksort(listaDocumentosOV, 0, listaDocumentosOV.size()-1);
		return listaDocumentosOV;
	}
	  
	  protected static void quicksort(List<CompleteDocuments> A, int izq, int der) {

		  CompleteDocuments pivote=A.get(izq); // tomamos primer elemento como pivote
		  int i=izq; // i realiza la búsqueda de izquierda a derecha
		  int j=der; // j realiza la búsqueda de derecha a izquierda
		  CompleteDocuments aux;
		 
		  while(i<j){            // mientras no se crucen las búsquedas
		     while(getIDOV(A.get(i))<=getIDOV(pivote) && i<j) i++; // busca elemento mayor que pivote
		     while(getIDOV(A.get(j))>getIDOV(pivote)) j--;         // busca elemento menor que pivote
		     if (i<j) {                      // si no se han cruzado                      
		         aux= A.get(i);                  // los intercambia
		         A.set(i, A.get(j));
		         A.set(j,aux);
		     }
		   }
		  A.set(izq,A.get(j)); // se coloca el pivote en su lugar de forma que tendremos
		  A.set(j,pivote); // los menores a su izquierda y los mayores a su derecha
		   if(izq<j-1)
		      quicksort(A,izq,j-1); // ordenamos subarray izquierdo
		   if(j+1 <der)
		      quicksort(A,j+1,der); // ordenamos subarray derecho
		}
	  
	  
	  private static Long getIDOV(CompleteDocuments completeDocuments) {
			Long IDOV=completeDocuments.getClavilenoid();
			for (CompleteElement elemetpos : completeDocuments.getDescription()) {
				if (elemetpos instanceof CompleteTextElement&&elemetpos.getHastype() instanceof CompleteTextElementType&&StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)elemetpos.getHastype()))
					{
					String IDOVS=((CompleteTextElement) elemetpos).getValue();
					try {
						
						IDOV=Long.parseLong(IDOVS);
						return IDOV;
					} catch (Exception e) {
						System.err.println("error de un entero que es "+IDOVS);
						e.printStackTrace();
					
					}
					}
			}
			return IDOV;
			
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
			
//			HashMap<Long, Integer> clave, 
			
			CompleteLogAndUpdates cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura, Long gramarId) {
		  
	  
		HashMap<Long, Integer> clave=new HashMap<Long, Integer>();
		List<CompleteElementType> ListaElementos;
	     
		CompleteTextElementType IDOV=null;
		
		if (grammar!=null)
	        ListaElementos=generaLista(grammar);
		else ListaElementos=new ArrayList<CompleteElementType>();
	        

			for (CompleteElementType completeElementType : ListaElementos) {
			
				if (completeElementType instanceof CompleteTextElementType && StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)completeElementType))
					IDOV=(CompleteTextElementType)completeElementType;
			}
		
			
			//Quito el IDOV de la lista
			if (IDOV!=null)
				ListaElementos.remove(IDOV);
			
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
		            		Value="Id Objeto Virtual ( USAR NUMEROS NEGATIVOS PARA AÑADIR NUEVOS ELEMENTOS )";
		            	else 
		            		if (j==1)
		            			Value="Descripción";
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
		            		Value="Identificador tipo( NO MODIFICAR )";
		            	else 
		            		if (j==1)
		            			Value="--";
		            		else
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);
		            		
		            		Integer I=StaticFuctionsOdAaXLS.getIDODAD(TmpEle);
		            		
		            		if (I!=null)
		            			Value=Integer.toString(I);
		            		else
		            			Value="#"+Long.toString(TmpEle.getClavilenoid());
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
	            
	            String DocID=null;
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					
					if (val==null&&elem instanceof CompleteTextElement && elem.getHastype() instanceof CompleteTextElementType &&StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)elem.getHastype()))
						DocID=((CompleteTextElement)elem).getValue();
					

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
	            	{	
	            		if (DocID!=null)
	            			Value=DocID;
	            		else
	            			Value="#"+Long.toString(Doc.getClavilenoid());
	            	}
	            	else if (c==1)
	            		Value=Doc.getDescriptionText();
	            	else
	            		{
	            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
	            		if (temp!=null)
	            		{
	            			if (temp.size()>0){
	            			CompleteElement completeElement=temp.get(0);
	            			Value=getValueFromElement(completeElement,cL);
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
//			HashMap<Long, Integer> clave,
			CompleteLogAndUpdates cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  

		
		HashMap<Long, Integer> clave=new HashMap<Long, Integer>();
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
		            		Value="Identificador Objeto Virtual";
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
		            		Value="Identificador tipo( NO MODIFICAR )";
		            	else 
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-1);

		            		Integer I=StaticFuctionsOdAaXLS.getIDODAD(TmpEle);
		            		
		            		if (I!=null)
		            			Value=Integer.toString(I);
		            		else
		            			Value="#"+Long.toString(TmpEle.getClavilenoid());
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
	            
	            String DocID=null;
	            
				for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					
					if (val==null&&elem instanceof CompleteTextElement && elem.getHastype() instanceof CompleteTextElementType &&StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)elem.getHastype()))
						DocID=((CompleteTextElement)elem).getValue();
					
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
	            	{
	            		if (DocID!=null)
	            			Value=DocID;
	            		else
	            			Value="#"+Long.toString(Doc.getClavilenoid());
	       
	            	}
	            	else
	            		{
	            		{
		            		ArrayList<CompleteElement> temp = ListaClave.get(c-1);
		            		if (temp!=null)
		            		{
		            			if (temp.size()>0){
		            			CompleteElement completeElement=temp.get(0);
		            			Value=getValueFromElement(completeElement,cL);
		            			
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
//			HashMap<Long, Integer> clave,
			CompleteLogAndUpdates cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  

		
		HashMap<Long, Integer> clave=new HashMap<Long, Integer>();
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
		            		Value=" Identificador Objeto Virtual dueño";
		            	else if (j==1)
		            		Value="Identificador referencia (Identificador del Objeto Virtual en Datos/Metadatos, de los recursos en Archivos/Url)";
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
		            		Value="Identificador tipo( NO MODIFICAR )";
		        		else if (j==1)
		            		Value=Long.toString(grammar.getClavilenoid());
		            	else 
		            		{
		            		CompleteElementType TmpEle = ListaElementos.get(j-2);

		            		Integer I=StaticFuctionsOdAaXLS.getIDODAD(TmpEle);
		            		
		            		if (I!=null)
		            			Value=Integer.toString(I);
		            		else
		            			Value="#"+Long.toString(TmpEle.getClavilenoid());
		            		
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
	            HashMap<Integer,String> ListaRelacionesporAmbito=new HashMap<Integer,String>();
	            
	            String DocID=null;
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					
					if (val==null&&elem instanceof CompleteTextElement && elem.getHastype() instanceof CompleteTextElementType &&StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)elem.getHastype()))
						DocID=((CompleteTextElement)elem).getValue();
						
					
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
						if (((CompleteLinkElement) elem).getValue() != null)
						{
							String DocID2=null;
							for (CompleteElement elem2 : ((CompleteLinkElement) elem).getValue().getDescription()) {	
								if (elem2 instanceof CompleteTextElement && elem2.getHastype() instanceof CompleteTextElementType &&StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)elem2.getHastype()))
										DocID2=((CompleteTextElement)elem2).getValue();
												
							}
						
							String Val=null;	
						
						if (DocID2!=null)
						{
							try {
								Val=DocID2;	
							} catch (Exception e) {
								Val ="#"+Long.toString(((CompleteLinkElement) elem).getValue().getClavilenoid());
							}
						}else	
							{
							Val ="#"+Long.toString(((CompleteLinkElement) elem).getValue().getClavilenoid());	
							}
						
						ListaRelacionesporAmbito.put(elem.getAmbitos().get(0), Val);
						
						}
					}
				}
	            
	            
	            
	            
	            for (Entry<Integer, String> listarelaciones : ListaRelacionesporAmbito.entrySet()) {
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
		            		{
		            		if (DocID!=null)
		            			Value=DocID;
		            		else
		            			Value="#"+Long.toString(Doc.getClavilenoid());
		            		}
		            	else if (c==1)
		            		Value=listarelaciones.getValue();
		            	else
		            		{
		            		{
			            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
			            		if (temp!=null)
			            		{
			            			if (temp.size()>0){
			            			CompleteElement completeElement=temp.get(0);
			            			Value=getValueFromElement(completeElement,cL);
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
//			HashMap<Long, Integer> clave,
			CompleteLogAndUpdates cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  
		HashMap<Long, Integer> clave=new HashMap<Long, Integer>();
	        
		  
			List<CompleteElementType> ListaElementos=new ArrayList<CompleteElementType>();
			
			CompleteElementType Filefisico=findFilesFisico(grammar.getSons());
			CompleteElementType FileOwner=findOwner(grammar.getSons());

			
			if (Filefisico==null)
				{
				cL.getLogLines().add("No posee Path de los archivos");
				Filefisico=new CompleteElementType();
				Filefisico.setClavilenoid(-1l);
				}
			
			if (FileOwner==null)
			{
				cL.getLogLines().add("No posee dueño de los archivos");
				FileOwner=new CompleteElementType();
				FileOwner.setClavilenoid(-2l);
			}
			
			
			Filefisico.setName("Path archivo");
			FileOwner.setName("Identificador del Objeto Virtual dueño");
			
			ListaElementos.add(FileOwner);
			ListaElementos.add(Filefisico);
			
			
			
	      
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        {
	        

	        	Row fila = hoja.createRow(row);
	        	row++;

	        	
	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Identificador archivo";
		            	else if (j==1)
		            		Value="Descripción";
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
	            		Value="#"+Long.toString(Doc.getClavilenoid());
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
		            			Value=getValueFromElement(completeElement,cL);
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
//			HashMap<Long, Integer> clave,
			CompleteLogAndUpdates cL, List<CompleteDocuments> ListaDocumentos, boolean soloEstructura) {
		  
		HashMap<Long, Integer> clave=new HashMap<Long, Integer>();
	        
		  
			List<CompleteElementType> ListaElementos=new ArrayList<CompleteElementType>();
			
			CompleteElementType URi=findURI(grammar.getSons());

			if (URi==null)
			{
			cL.getLogLines().add("No posee URI de las URLs");
			URi=new CompleteElementType();
			URi.setClavilenoid(-1l);
			}
			
			URi.setName("URI");
			
			ListaElementos.add(URi);
	      
	        if (ListaDocumentos.size()+2>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls solo se exportaran los 65534 primeros");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65534);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        {
	        		Row fila = hoja.createRow(row);
	        		row++;
	        	

	        		for (int j = 0; j < columnsMax+2; j++) {
		        		
		        		String Value = "";
		            	if (j==0)
		            		Value="Identificadotr URL";
		            	else if (j==1)
		            		Value="Descripción";
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
	            		Value="#"+Long.toString(Doc.getClavilenoid());
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
		            			Value=getValueFromElement(completeElement,cL);
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
			 		if ((completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)&&(!StaticFuctionsOdAaXLS.isIgnored((CompleteElementType)completeelem)))
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
		 
		  
		  
		  processCompleteCollection(new CompleteLogAndUpdates(), CC, false, System.getProperty("user.home"));
		  
	    }


	private static String getValueFromElement(CompleteElement completeElement,CompleteLogAndUpdates cL) {
		try {
			if (completeElement instanceof CompleteTextElement)
    			{
				String ValueText=((CompleteTextElement)completeElement).getValue();
				if (StaticFuctionsOdAaXLS.isNumeric(completeElement.getHastype()))
				{
					try {
						Double D=Double.parseDouble(ValueText);
						ValueText=D.toString();
						
						int p_ent= (int)D.intValue();
						 
						double p_dec= D - p_ent;
						
						if (p_dec==0)
							ValueText=Integer.toString(p_ent);
					
					} catch (Exception e2) {
					}
				}
				if (StaticFuctionsOdAaXLS.isDate(completeElement.getHastype()))
				{
					try {

						
						Date fecha = null;
						//yyyy-MM-dd HH:mm:ss
						try {
							SimpleDateFormat formatoDelTexto = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
							fecha = formatoDelTexto.parse(ValueText);
						} catch (Exception e) {
							//Nada
							fecha = null;
						}
						
						if (fecha==null)
							try {
								SimpleDateFormat formatoDelTexto = new SimpleDateFormat("yyyy-MM-dd");
								fecha = formatoDelTexto.parse(ValueText);
							} catch (Exception e) {
								//Nada
								fecha = null;
							}
						
						if (fecha==null)
							try {
								SimpleDateFormat formatoDelTexto = new SimpleDateFormat("dd-MM-yyyy");
								fecha = formatoDelTexto.parse(ValueText);
							} catch (Exception e) {
								//Nada
								fecha = null;
							}
						
						if (fecha==null)
							try {
								SimpleDateFormat formatoDelTexto = new SimpleDateFormat("yyyy-MM-dd HH:mm");
								fecha = formatoDelTexto.parse(ValueText);
							} catch (Exception e) {
								//Nada
								fecha = null;
							}
						
						if (fecha==null)
							try {
								SimpleDateFormat formatoDelTexto = new SimpleDateFormat("yyyyMMdd");
								fecha = formatoDelTexto.parse(ValueText);
							} catch (Exception e) {
								//Nada
								fecha = null;
							}
						
						if (fecha==null)
							try {
								SimpleDateFormat formatoDelTexto = new SimpleDateFormat("dd/MM/yyyy");
								fecha = formatoDelTexto.parse(ValueText);
							} catch (Exception e) {
								//Nada
								fecha = null;
							}
						
						if (fecha==null)
							try {
								SimpleDateFormat formatoDelTexto = new SimpleDateFormat("dd/MM/yy");
								fecha = formatoDelTexto.parse(ValueText);
							} catch (Exception e) {
								//Nada
								fecha = null;
							}
						
						if (fecha!=null)
						{
						DateFormat df = new SimpleDateFormat ("dd/MM/yyyy");
						ValueText=df.format(fecha);	
						}
						else
						{
							cL.getLogLines().add("Problemas al parsear la fecha  " + ValueText + "  solo formatos compatibles yyyy-MM-dd HH:mm:ss ó yyyy-MM-dd HH:mm ó yyyy-MM-dd ó yyyyMMdd ó dd/MM/yyyy ó dd/MM/yy ó dd-MM-yyyy");
						}
					} catch (Exception e2) {
					}

				}
				return ValueText;
    			}
			else if (completeElement instanceof CompleteLinkElement)
				{
				CompleteDocuments Elem = ((CompleteLinkElement)completeElement).getValue();
				if (Elem!=null)
					{
					String Value = getIdov(Elem);
					if (Value!=null)
						return Value;
					else
						return "#"+Elem.getClavilenoid();
//					return Long.toString((((CompleteLinkElement)completeElement).getValue().getClavilenoid()));
					}
				return "";
				}
			else if (completeElement instanceof CompleteResourceElementURL)
				return (((CompleteResourceElementURL)completeElement).getValue());
			else if (completeElement instanceof CompleteResourceElementFile)
				return (((CompleteResourceElementFile)completeElement).getValue().getPath());
		} catch (Exception e) {
			return "";
		}
		return "";
	}
	
	private static String getIdov(CompleteDocuments Doc) {
		for (CompleteElement elem : Doc.getDescription())
			
			if (elem instanceof CompleteTextElement && elem.getHastype() instanceof CompleteTextElementType &&StaticFuctionsOdAaXLS.isIDOV((CompleteTextElementType)elem.getHastype()))
				return ((CompleteTextElement)elem).getValue();
		
		return null;
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
