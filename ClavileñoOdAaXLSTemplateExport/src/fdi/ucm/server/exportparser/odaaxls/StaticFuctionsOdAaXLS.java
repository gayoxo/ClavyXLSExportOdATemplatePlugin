/**
 * 
 */
package fdi.ucm.server.exportparser.odaaxls;

import java.util.ArrayList;

import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalValueType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalView;

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
	 * Revisa si un elemento es Recursos
	 * @param hastype
	 * @return
	 */
	public static boolean isRecursos(CompleteElementType hastype) {
		
		ArrayList<CompleteOperationalView> Shows = hastype.getShows();
		for (CompleteOperationalView show : Shows) {
			
			if (show.getName().equals(StaticNamesOdAaXLS.META))
			{
				ArrayList<CompleteOperationalValueType> ShowValue = show.getValues();
				for (CompleteOperationalValueType CompleteOperationalValueType : ShowValue) {
					if (CompleteOperationalValueType.getName().equals(StaticNamesOdAaXLS.TYPE))
						if (CompleteOperationalValueType.getDefault().equals(StaticNamesOdAaXLS.RECURSO)) 
										return true;
				}
			}
		}
		return false;
	}
}
