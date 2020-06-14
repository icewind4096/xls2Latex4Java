import com.mindmotion.xls2latex.convert.Arg2Paramater;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.file.GeneralTabFile;
import com.mindmotion.xls2latex.file.RegTabFile;
import com.mindmotion.xls2latex.util.FileUtil;

public class XLS2Latex {
    public static void main(String[] args) {
        ParamaterInfo paramaterInfo = Arg2Paramater.arg2Paramater(args);

        if (FileUtil.FileExists(paramaterInfo.getSourceFileName()) == false) {
            System.exit(ResultEnum.EXCELFILENOTEXIST.getCode());
        };

        int resultCode = 0;
        if (paramaterInfo.getType() == 1){
            resultCode = RegTabFile.GenerateFile(paramaterInfo);
        } else {
            resultCode = GeneralTabFile.GenerateFile(paramaterInfo);
        }

        if (resultCode != ResultEnum.SUCCESS.getCode()) {
            System.exit(resultCode);
        }
    }
}
