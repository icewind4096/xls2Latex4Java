import com.mindmotion.xls2latex.convert.Arg2Paramater;
import com.mindmotion.xls2latex.domain.ParamaterInfo;
import com.mindmotion.xls2latex.enums.GeneralFileTypeEnum;
import com.mindmotion.xls2latex.enums.ResultEnum;
import com.mindmotion.xls2latex.file.GeneralTabFile;
import com.mindmotion.xls2latex.file.RegTabFile;
import com.mindmotion.xls2latex.util.FileUtil;

public class XLS2Latex {
    public static void main(String[] args) {
        //type -> 2  D:\excel\latex\cn_FlashOrganization.tex D:\excel\cn_FlashOrganization.xlsx 2 430 4 70,90,140,140 0 2 FLASH模块操作寄存器一览表
        //type -> 2  D:\excel\latex\cn_Replace.tex D:\excel\cn_Replace.xlsx 0 430 5 70,50,70,140,100 0 2 FLASH模块操作寄存器一览表
        //type -> 1  D:\excel\latex D:\excel\cn_FLASH.xlsx 1 440 5 90,90,90,170,90 0 0
        //type -> 0  D:\excel\latex\cn_FLASH_OverView.tex D:\excel\cn_FLASH_OverView.xlsx 0 430 5 70,50,70,140,100 0 2 FLASH模块操作寄存器一览表
        ParamaterInfo paramaterInfo = Arg2Paramater.arg2Paramater(args);

        if (FileUtil.FileExists(paramaterInfo.getSourceFileName()) == false) {
            System.exit(ResultEnum.EXCELFILENOTEXIST.getCode());
        };

        int resultCode = 0;
        if (paramaterInfo.getGeneralFileTypeEnum() == GeneralFileTypeEnum.REGFILE){
            resultCode = RegTabFile.GenerateFile(paramaterInfo);
        } else {
            resultCode = GeneralTabFile.GenerateFile(paramaterInfo);
        }

        if (resultCode != ResultEnum.SUCCESS.getCode()) {
            System.exit(resultCode);
        }
    }
}
