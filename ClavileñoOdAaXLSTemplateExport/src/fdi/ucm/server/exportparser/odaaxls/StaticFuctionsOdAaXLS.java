/**
 * 
 */
package fdi.ucm.server.exportparser.odaaxls;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteOperationalValue;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalValueType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

/**
 * Funcion que implementa las funciones estaticas de la exportacion
 * @author Joaquin Gayoso-Cabada
 *
 */
public class StaticFuctionsOdAaXLS {
	
	/**
	 * Revisa si un elemento es VirtualObject
	 * @param hastype
	 * @return
	 */
	public static boolean isVirtualObject(CompleteGrammar hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getViews();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{
					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.VIRTUAL_OBJECT)) 
										return true;

			}
		}
		return false;
	}
	
	/**
	 * Revisa si un elemento es VirtualObject
	 * @param hastype
	 * @return
	 */
	public static boolean isFiles(CompleteGrammar hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getViews();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.FILE)) 
										return true;

			}
		}
		return false;
	}
	
	
	
	/**
	 * Revisa si un elemento es FileResource
	 * @param hastype
	 * @return
	 */
	public static boolean isFileFisico(CompleteElementType hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.FILERESOURCE)) 
										return true;

			}
		}
		return false;
	}
	
	/**
	 * Revisa si un elemento es Owner
	 * @param hastype
	 * @return
	 */
	public static boolean isOwner(CompleteElementType hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.IDOV_OWNER)) 
										return true;

			}
		}
		return false;
	}
	
	
	/**
	 * Revisa si un elemento es METADATOS
	 * @param hastype
	 * @return
	 */
	public static boolean isDatos(CompleteElementType hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.DATOS)) 
										return true;

			}
		}
		return false;
	}

	
	/**
	 * Revisa si un elemento es METADATOS
	 * @param hastype
	 * @return
	 */
	public static boolean isMetaDatos(CompleteElementType hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.METADATOS)) 
										return true;

			}
		}
		return false;
	}
	

	/**
	 * Revisa si un elemento es Resources
	 * @param hastype
	 * @return
	 */
	public static boolean isResources(CompleteElementType hastype) {
		
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.RESOURCE)) 
										return true;

			}
		}
		return false;
	}

	public static boolean isURL(CompleteGrammar completeGrammar) {
		ArrayList<CompleteOperationalValueType> Shows = completeGrammar.getViews();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.URL)) 
										return true;

			}
		}
		return false;
	}

	public static boolean isURI(CompleteElementType completeStructure) {
		ArrayList<CompleteOperationalValueType> Shows = completeStructure.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.URI)) 
										return true;

			}
		}
		return false;
	}
	
	public static Integer getIDODAD(CompleteElementType attribute) {
		ArrayList<CompleteOperationalValueType> Shows = attribute.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.PRESNTACION))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.OdaID))
						try {
							Integer I=Integer.parseInt(show.getDefault());
								return I;
						} catch (Exception e) {
							return null;
						}

			}
		}
		return null;
		
	}
	
	public static boolean isIDOV(CompleteTextElementType hastype) {
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.IDOV)) 
										return true;

			}
		}
		return false;
	}
	

	/**
	 * Clase que define si es numerico
	 * @param hastype
	 * @return
	 */
	public static boolean isNumeric(CompleteElementType hastype) {
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {	
			if (show.getView().equals(StaticNamesOdAaXLS.METATYPE))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.METATYPETYPE))
							if (show.getDefault().equals(StaticNamesOdAaXLS.NUMERIC)) 
										return true;
			}
		}
		return false;
	}








	public static boolean isNumeric(CompleteElement elem) {
		ArrayList<CompleteOperationalValue> ShowsInst = elem.getShows();
		for (CompleteOperationalValue show : ShowsInst) {
			if (show.getType().getView().equals(StaticNamesOdAaXLS.METATYPE)&&show.getType().getName().equals(StaticNamesOdAaXLS.METATYPETYPE))
				if (show.getValue().equals(StaticNamesOdAaXLS.NUMERIC))
					return true;
		}
		
		return isNumeric(elem.getHastype());
		
	}








	public static boolean isDate(CompleteElementType attribute) {
		ArrayList<CompleteOperationalValueType> Shows = attribute.getShows();
		for (CompleteOperationalValueType show : Shows) {	
			if (show.getView().equals(StaticNamesOdAaXLS.METATYPE))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.METATYPETYPE))
							if (show.getDefault().equals(StaticNamesOdAaXLS.DATE)) 
										return true;

			}
		}
		return false;
	}
	
	public static boolean isIgnored(CompleteElementType hastype) {
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(StaticNamesOdAaXLS.META))
			{

					if (show.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (show.getDefault().equals(StaticNamesOdAaXLS.IGNORED)) 
										return true;

			}
		}
		return false;
	}
	
	public static boolean isInGrammar(CompleteDocuments iterable_element,
			CompleteGrammar completeGrammar) {
		HashSet<Long> ElemT=new HashSet<Long>();
		for (CompleteElement dd : iterable_element.getDescription()) {
			ElemT.add(dd.getHastype().getClavilenoid());
		}
		
		return isInGrammar(ElemT, completeGrammar.getSons());
		
		
	}



	private static boolean isInGrammar(HashSet<Long> elemT,
			List<CompleteElementType> sons) {
		for (CompleteElementType CSlong1 : sons) {
			if (elemT.contains(CSlong1.getClavilenoid())||isInGrammar(elemT, CSlong1.getSons()))
				return true;
			
		}
		return false;
	}
}
