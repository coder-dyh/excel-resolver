package excel;

import java.math.BigDecimal;
import java.sql.Date;
import java.sql.SQLException;
import java.sql.Timestamp;

/**
 * 从数据库中取出对应的数据，这里要进行判断的原因是因为从数据库中取出的数据并不知道是什么数据类型
 */
public class TypeConvert {

    public static Object convert(Class type, Object value) throws SQLException {
        if (value == null) {
            throw new RuntimeException("值为null，无法转换");
        }
        if (type.equals(String.class)) {
            value = String.valueOf(value);
        } else if (type.equals(Integer.TYPE) || type.equals(Integer.class)) {
            value = Integer.valueOf(String.valueOf(value));
        } else if (type.equals(Double.TYPE) || type.equals(Double.class)) {
            value = Integer.valueOf(String.valueOf(value));
        } else if (type.equals(Long.TYPE) || type.equals(Long.class)) {
            value = Long.valueOf(String.valueOf(value));
        } else if (type.equals(Short.TYPE) || type.equals(Short.class)) {
            value = Short.valueOf(String.valueOf(value));
        } else if (type.equals(Byte.TYPE) || type.equals(Byte.class)) {
            value = Byte.valueOf(String.valueOf(value));
        } else if (type.equals(Float.TYPE) || type.equals(Float.class)) {
            value = Float.valueOf(String.valueOf(value));
        } else if (type.equals(BigDecimal.class)) {
            value = new BigDecimal(String.valueOf(value));
        } else if (type.equals(Character.TYPE) || type.equals(Character.class)) {
            value = Float.valueOf(String.valueOf(value));
        } else {
            try {
                dateConvert(value, type);
            } catch (Exception e) {
                throw new RuntimeException("为属性设置值时，无法解析的属性类型，暂不支持对象类型、Boolean类型...");
            }
        }

        return value;
    }

    /**
     * 时间转换工具
     *
     * @param value
     * @param type
     * @return
     */
    private static Object dateConvert(Object value, Class<?> type) {
        if (type.equals(Date.class)) {
            value = new Date(((java.util.Date) value).getTime());
        } else if (type.equals(java.sql.Time.class)) {
            value = new java.sql.Time(((java.util.Date) value).getTime());
        } else {
            Timestamp ts = (Timestamp) value;
            //获得时间戳的最小精度
            int nanos = ts.getNanos();
            value = new Timestamp(ts.getTime());
            ts.setNanos(nanos);
        }
        return value;
    }
}