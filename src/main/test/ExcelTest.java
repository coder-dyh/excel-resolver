import excel.ExcelHelper;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

public class ExcelTest {

    public static void main(String[] args) throws Exception {
        InputStream is = new FileInputStream("/Users/yooee/Downloads/对账模板/放款/富友格式.xlsx");
        List<BalanceAccountDto> list = ExcelHelper.parseExcelValueToBean(BalanceAccountDto.class, is);
        list.forEach(a -> {
            System.out.println("amount:" + a.getAmount() + "    orderId:" + a.getOrderId());
        });
    }

}
