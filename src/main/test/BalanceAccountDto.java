import excel.annotation.ExcelValue;

import java.io.Serializable;
import java.math.BigDecimal;

public class BalanceAccountDto implements Serializable{
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@ExcelValue(sort = {"5", "15"}, value = "15")
	private String orderId;
	@ExcelValue(sort = {"15", "5"}, value = "5")
	private BigDecimal amount;
	public String getOrderId() {
		return orderId;
	}
	public void setOrderId(String orderId) {
		this.orderId = orderId;
	}
	public BigDecimal getAmount() {
		return amount;
	}
	public void setAmount(BigDecimal amount) {
		this.amount = amount;
	}
	
}
