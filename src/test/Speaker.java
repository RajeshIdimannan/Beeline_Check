package test;

import java.io.File;

import javax.sound.sampled.AudioFormat;
import javax.sound.sampled.AudioInputStream;
import javax.sound.sampled.AudioSystem;
import javax.sound.sampled.DataLine;
import javax.sound.sampled.SourceDataLine;

public class Speaker {

	
	 File soundFile;

	    public Speaker(String file) {
	        soundFile = new File(file);
	    }

	    public void play() {
	        try {
	            // create audio input stream to file
	            AudioInputStream ais = AudioSystem.getAudioInputStream(
	                soundFile);
	            // determine the file's audio format
	            AudioFormat format = ais.getFormat();
	            System.out.println("Format: " + format);
	            // get a line to play the audio
	            DataLine.Info info = new DataLine.Info(
	                SourceDataLine.class, format);
	            SourceDataLine source = (SourceDataLine) AudioSystem.getLine(
	                info);
	            // play the file
	            source.open(format);
	            source.start();
	            int read = 0;
	            byte[] audioData = new byte[16384];
	            while (read > -1) {
	                read = ais.read(audioData, 0, audioData.length);
	                if (read >= 0) {
	                    source.write(audioData, 0, read);
	                }
	            }
	            source.drain();
	            source.close();
	        } catch (Exception exc) {
	            System.out.println("Error: " + exc.getMessage());
	            exc.printStackTrace();
	        }
	        System.exit(0);
	    }
	public static void main(String[] arguments) {
		if (arguments.length < 1) {
            System.out.println("Usage: java Speaker filename");
            System.exit(-1);
        }
        Speaker speaker = new Speaker(arguments[0]);
        speaker.play();

	}

}
