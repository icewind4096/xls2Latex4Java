import com.mindmotion.xls2latex.convert.Arg2Paramater;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.file.ExcelFile;
import com.mindmotion.xls2latex.util.FileUtil;

public class xls2Latex {
    public static void main(String[] args) {
        ParamaterInfo paramaterInfo = Arg2Paramater.arg2Paramater(args);

        if (FileUtil.FileExists(paramaterInfo.getSourceFileName()) == false) {
            System.exit(ResultEnum.EXCELFILENOTEXIST.getCode());
        };

        if (paramaterInfo.getType() == 1){
            int resultCode = ExcelFile.ProduceRegTab(paramaterInfo);

            if (resultCode != ResultEnum.SUCCESS.getCode()) {
                System.exit(resultCode);
            }
        }
    }
}
