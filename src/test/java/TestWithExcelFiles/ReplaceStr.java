package TestWithExcelFiles;

public class ReplaceStr {

    public static String replaceString(String before) {
        return before.replaceAll("<CONDITION type=\"SINGULAR_CONCRETE\">", "").replaceAll("<CONDITION>", "").replaceAll("<COLUMN>", "").replaceAll("<ALIAS>", " ").replaceAll("</COLUMN>", "").replaceAll("</ALIAS>", ".").replaceAll("</CONDITION>", "").replaceAll("<OPERATION>", "").replaceAll("</OPERATION>", " ").replaceAll("<COPULA>", " ").replaceAll("</COPULA>", "").replaceAll("<COLNAME>", "").replaceAll("</COLNAME>", " ").replaceAll("<CONDITION type=\"EXPRESSION_FREEHAND\">","").replaceAll("<OPERATION/>"," ").replaceAll("<VALUE>","").replaceAll("</VALUE>","").replaceAll("<OPERATION>"," ").replaceAll("<OPERATION/>","").replaceAll("<VALUE/>","").replaceAll("<CONDITION type=\"EXPRESSION_SINGULAR\">","").replaceAll("<CONDITION type=\"EXPRESSION_SHORTHAND\">","").replaceAll("\\n","").replaceAll("<CONDITION type=\"EXTERNAL_LIST\">","");
    }
    public static String nameOfField (String condition){
        return condition.substring(condition.indexOf("<COLNAME>") + 9,condition.indexOf("</COLNAME>"));
    }
}
