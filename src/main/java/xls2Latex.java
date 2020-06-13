import com.mindmotion.xls2Latex.convert.Arg2Paramater;
import com.mindmotion.xls2Latex.domain.ParamaterInfo;
import com.mindmotion.xls2Latex.file.ExcelFile;
import com.mindmotion.xls2Latex.util.FileUtil;

public class xls2Latex {
    public static void main(String[] args) {
        ParamaterInfo paramaterInfo = Arg2Paramater.arg2Paramater(args);

        if (FileUtil.FileExists(paramaterInfo.getSourceFileName()) == false) {
            System.exit(3);
        };

        if (paramaterInfo.getType() == 1){
            ExcelFile.ProduceRegTab(paramaterInfo);
        }
    }
}
