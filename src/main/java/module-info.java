module com.lifeinsuranceapplication {
    requires javafx.controls;
    requires javafx.fxml;
    requires java.sql;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;


    opens com.lifeinsuranceapplication to javafx.fxml;
    exports com.lifeinsuranceapplication;
}