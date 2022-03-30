/**
 * 
 */
package GoogleSheetProject;

/**
 * @author yuval
 *Yuval Mastey
 * Tamir Spilberg
 */
public class ExceptionRowWasNotFound extends Exception {
	public ExceptionRowWasNotFound(String id) {
		super("ERROR! Row with id "+id+" wans't found");
	}
}
