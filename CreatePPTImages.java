import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class CreatePPTImages {
	public static void main(String args[]) throws IOException {
		// creating an empty presentation
		String inputFileLocation = "C://testFolder//test.pptx";
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(new File(
				inputFileLocation)));
		// getting the dimensions and size of the slide
		Dimension pgsize = ppt.getPageSize();
		XSLFSlide[] slide = ppt.getSlides();
		// creating image from the slide one by one
		for (int i = 0; i < slide.length; i++) {
			BufferedImage image = new BufferedImage(pgsize.width,
					pgsize.height, BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = image.createGraphics();
			graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
					RenderingHints.VALUE_ANTIALIAS_ON);
			graphics.setRenderingHint(RenderingHints.KEY_RENDERING,
					RenderingHints.VALUE_RENDER_QUALITY);
			graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
					RenderingHints.VALUE_INTERPOLATION_BICUBIC);
			graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,
					RenderingHints.VALUE_FRACTIONALMETRICS_ON);
			// clear the drawing area
			graphics.setPaint(Color.white);
			graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width,
					pgsize.height));
			// render slide
			slide[i].draw(graphics);
			// creating an image file as output
			FileOutputStream out = new FileOutputStream("ppt_image" + i
					+ ".png");
			javax.imageio.ImageIO.write(image, "png", out);
			ppt.write(out);
			out.close();
		}
	}
}
