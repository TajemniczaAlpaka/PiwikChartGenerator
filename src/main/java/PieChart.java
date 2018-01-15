import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PiePlot3D;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.util.Rotation;

import javax.swing.*;
import java.util.ArrayList;
import java.util.Collections;

public class PieChart extends JFrame {

    private static final long serialVersionUID = 1L;

    public PieChart(String applicationTitle, String chartTitle, ArrayList<String> data) {
        super(applicationTitle);
        // This will create the dataset
        PieDataset dataset = createDataset(data);
        // based on the dataset we create the chart
        JFreeChart chart = createChart(dataset, chartTitle);
        // we put the chart into a panel
        ChartPanel chartPanel = new ChartPanel(chart);
        // default size
        chartPanel.setPreferredSize(new java.awt.Dimension(500, 270));
        // add it to our application
        setContentPane(chartPanel);

        this.setExtendedState(JFrame.MAXIMIZED_BOTH);

    }


    public PieDataset createDataset(ArrayList<String> data) {

        DefaultPieDataset result = new DefaultPieDataset();
        ArrayList<String> alreadyOccurred = new ArrayList<>();

        for (String value : data) {
            if (!alreadyOccurred.contains(value) && !value.equals("\\N")) {
                alreadyOccurred.add(value);
                int occurrences = Collections.frequency(data, value);
                result.setValue(value, occurrences);
            }
        }
        return result;
    }

    /**
     * Creates a chart
     */
    private JFreeChart createChart(PieDataset dataset, String title) {

        JFreeChart chart = ChartFactory.createPieChart(
                title,                  // chart title
                dataset,                // data
                true,                   // include legend
                true,
                false
        );

        PiePlot plot = (PiePlot) chart.getPlot();
        plot.setStartAngle(290);
        plot.setDirection(Rotation.CLOCKWISE);
        plot.setForegroundAlpha(0.5f);
        return chart;

    }
}