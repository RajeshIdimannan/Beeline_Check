package BeelineData;
import static java.util.concurrent.TimeUnit.*;

import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;

public class TaskScheduler {		
	 
	    private final ScheduledExecutorService scheduler =
	       Executors.newScheduledThreadPool(1);

	    	public static void main(String []args){
	    		TaskScheduler tsk= new TaskScheduler();
	    		tsk.beepForAnHour();
	    	}
	    public  void beepForAnHour() {
	        final Runnable beeper = new Runnable() {
	                public void run() { 
	                	System.out.println("beep"); 
	                	}
	            };
	        final ScheduledFuture<?> beeperHandle =
	            scheduler.scheduleAtFixedRate(beeper, 5, 5, SECONDS);
	        scheduler.schedule(new Runnable() {
	                public void run() { 
	                	beeperHandle.cancel(true); 
	                }
	            }, 60 * 60, SECONDS);
	    }
	 }


