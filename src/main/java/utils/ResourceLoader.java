package utils;

import java.net.URL;
import java.util.concurrent.ExecutionException;
import javax.swing.SwingWorker;

public class ResourceLoader extends SwingWorker<URL, Object>{
    private String resourceName;

    public ResourceLoader(String resourceName) {
        this.resourceName = resourceName;
    }

    @Override
    protected URL doInBackground() throws Exception {
        return getClass().getClassLoader().getResource(resourceName);
    }

    @Override
    protected void done() {
        // use result here 
        try {
			URL result = get();
			System.out.println("--- CARREGADO ---");
		} catch (InterruptedException | ExecutionException e) {
			e.printStackTrace();
		}
    }
}
