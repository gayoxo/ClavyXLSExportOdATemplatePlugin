/**
 * 
 */
package fdi.ucm.server.exportparser.odaaxls;

import java.util.ArrayList;

import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteOperationalValue;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalValueType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalView;
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getViews();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.VIRTUAL_OBJECT)) 
										return true;

				}
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getViews();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.FILE)) 
										return true;

				}
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.FILERESOURCE)) 
										return true;

				}
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.IDOV_OWNER)) 
										return true;
				}
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.DATOS)) 
										return true;

				}
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.METADATOS)) 
										return true;

				}
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
		
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.RESOURCE)) 
										return true;
				}
			}
		}
		return false;
	}

	public static boolean isURL(CompleteGrammar completeGrammar) {
		ArrayList<CompleteOperationalView> Shows = completeGrammar.getViews();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.URL)) 
										return true;

				}
			}
		}
		return false;
	}

	public static boolean isURI(CompleteElementType completeStructure) {
		ArrayList<CompleteOperationalView> Shows = completeStructure.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.URI)) 
										return true;

				}
			}
		}
		return false;
	}
	
	public static Integer getIDODAD(CompleteElementType attribute) {
		ArrayList<CompleteOperationalView> Shows = attribute.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.PRESNTACION))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.OdaID))
						try {
							Integer I=Integer.parseInt(CompleteOperationalValueType.getDefault());
								return I;
						} catch (Exception e) {
							return null;
						}
						

				}
			}
		}
		return null;
		
	}
	
	public static boolean isIDOV(CompleteTextElementType hastype) {
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.IDOV)) 
										return true;

				}
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
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {	
			if (show.getName().equals(StaticNamesOdAaXLS.METATYPE))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType showValues : ShowValue) {
					if (showValues.getName().equals(StaticNamesOdAaXLS.METATYPETYPE))
							if (showValues.getDefault().equals(StaticNamesOdAaXLS.NUMERIC)) 
										return true;
				}
			}
		}
		return false;
	}








	public static boolean isNumeric(CompleteElement elem) {
		ArrayList<CompleteOperationalValue> ShowsInst = elem.getShows();
		for (CompleteOperationalValue show : ShowsInst) {
			if (show.getType().getView().getName().equals(StaticNamesOdAaXLS.METATYPE)&&show.getType().getName().equals(StaticNamesOdAaXLS.METATYPETYPE))
				if (show.getValue().equals(StaticNamesOdAaXLS.NUMERIC))
					return true;
		}
		
		return isNumeric(elem.getHastype());
		
	}








	public static boolean isDate(CompleteElementType attribute) {
		ArrayList<CompleteOperationalView> Shows = attribute.getShows();
		for (CompleteOperationalView show : Shows) {	
			if (show.getName().equals(StaticNamesOdAaXLS.METATYPE))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType showValues : ShowValue) {
					if (showValues.getName().equals(StaticNamesOdAaXLS.METATYPETYPE))
							if (showValues.getDefault().equals(StaticNamesOdAaXLS.DATE)) 
										return true;
				}
			}
		}
		return false;
	}
}
