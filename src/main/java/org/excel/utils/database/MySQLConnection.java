package org.excel.utils.database;

import java.sql.*;
import java.util.Properties;

/**
 * 继承实现
 */
public class MySQLConnection {

    private Connection connection;
    public MySQLConnection() throws ClassNotFoundException, SQLException {
        Class.forName("com.mysql.cj.jdbc.Driver");
        // 2.用户信息和url
        //String url = "jdbc:mysql://localhost:3306/shop?useUnicode=true&characterEncoding=utf8&useSSL=false";
        String url = "jdbc:mysql://localhost:3306/";
        Properties info = new Properties();
        info.setProperty("user","root");
        info.setProperty("password","43007884Ct!");
        info.setProperty("useUnicode","true");
        info.setProperty("characterEncoding","utf8");
        info.setProperty("useSSL","false");
        //String user = "root";
        //String password = "43007884Ct!";
        // 3.获取连接
        //connection = DriverManager.getConnection(url, user, password);
        connection = DriverManager.getConnection(url, info);

    }

    public void selectDB(String dbname) throws SQLException {
        PreparedStatement preparedStatement = connection.prepareStatement("use " + dbname + ";");
        preparedStatement.execute();
    }

    public ResultSet query(String sql) throws SQLException {
        return connection.prepareStatement(sql).executeQuery();
    }

    public void execute(String sql) throws SQLException {
        connection.prepareStatement(sql).execute();
    }
    public static void main(String[] args) throws SQLException, ClassNotFoundException {
        MySQLConnection connection1 = new MySQLConnection();
        connection1.selectDB("demo");
        ResultSet resultSet = connection1.query("select * from teacher where id=1");
        resultSet.next();
        String name = resultSet.getString(2);
        int stid = resultSet.getInt(3);
        System.out.println("name="+name+", student id="+stid);
    }

}
