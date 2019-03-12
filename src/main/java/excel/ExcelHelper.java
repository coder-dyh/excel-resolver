package excel;

import excel.annotation.ExcelValue;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @Discription
 * @Date 2019-03-11 21:01
 * @Author James
 */
public class ExcelHelper {

    private static Field[] fields;

    private static List<Field> list = new ArrayList<>();

    private static Map<Integer, Field> map = new TreeMap<>();

    private static List<Integer> locations = new ArrayList<>();

    /**
     * 通过注解 @ExcelValue 解析 excel 值到 bean 中
     * @param cls
     * @param is
     * @param isExcel2003
     */
    public static <T> List<T> parseExcelValueToBean(Class<?> cls, InputStream is, boolean isExcel2003) {
        return resolveExcel(cls, is, isExcel2003, null);
    }

    /**
     * 通过注解 @ExcelValue 解析 excel 值到 bean 中
     * @param cls
     * @param is
     */
    public static <T> List<T> parseExcelValueToBean(Class<?> cls, InputStream is) {
        return resolveExcel(cls, is, false, null);
    }

    /**
     * 通过注解 @ExcelValue 解析 excel 值到 bean 中
     * @param cls
     * @param is
     * @param isExcel2003
     * @param type 类型，使用 @ExcelValue(arr={1, 2})时指定需要选择哪种读取方式
     */
    public static <T> List<T> parseExcelValueToBean(Class<?> cls, InputStream is, boolean isExcel2003, Integer type) {
        return resolveExcel(cls, is, isExcel2003, type);
    }

    public static <T> List<T> resolveExcel(Class<?> cls, InputStream is, boolean isExcel2003, Integer type) {
        List<Integer> fieldList = getFieldWithExcelValueAnnotation(cls, type);
        List<String> values = ExcelUtil.readExcelValue(is, isExcel2003, fieldList.toArray());

        List<T> result = new ArrayList<>();
        // ExcelUtil.getTotalRows(is, isExcel2003)
        Object instance = null;
        for (int i = 0; i < values.size(); i++) {
            if (i % list.size() == 0) {
                try {
                    instance = cls.newInstance();
                    map.get(fieldList.get(i % list.size())).set(instance, TypeConvert.convert(map.get(fieldList.get(i % list.size())).getType(), values.get(i)));
                } catch (Exception e) {
                    e.printStackTrace();
                }
                result.add((T) instance);
            } else {
                try {
                    map.get(fieldList.get(i % list.size())).set(instance, TypeConvert.convert(map.get(fieldList.get(i % list.size())).getType(), values.get(i)));
                } catch (Exception e) {
                    throw new RuntimeException("解析后为属性设置值异常，请检查属性值和解析的 excel 后得到的值是否同一类型");
                }
            }

        }
        return result;
    }

    /**
     * 根据 Class 对象获取该对象中所有带有 ExcelValue 注解的属性
     * @param cls
     * @return List<Field>
     */
    private static List<Integer> getFieldWithExcelValueAnnotation(Class<?> cls, Integer type) {
        List<Integer> fieldList = new ArrayList<>();
        fields = cls.getDeclaredFields();
        for (Field f : fields) {
            if (f.getAnnotation(ExcelValue.class) != null) {
                Integer location = null;
                try {
                    if (type == null || type == 0) {
                        location = getExcelValueAnnotationValue(f);
                    } else {
                        location = getExcelValueAnnotationValue(type, f);
                    }
                    if (location == null) {
                        throw new RuntimeException("ExcelVaule 注解的值非法，必须是 Integer 类型");
                    }
                } catch (NumberFormatException e) {
                    throw new RuntimeException("ExcelVaule 注解的值非法，必须是 Integer 类型");
                }
                f.setAccessible(true);
                list.add(f);
                map.put(location, f);
                fieldList.add(location);
            }
            Collections.sort(fieldList);
        }
        return fieldList;
    }

    private static Integer getExcelValueAnnotationValue(int type, Field f) {
        Integer location = null;
        String[] val = f.getAnnotation(ExcelValue.class).sort();
        for (int i = 1; i <= val.length; i++) {
            if (type == i) {
                location = Integer.valueOf(val[i]);
            }
        }
        return location;
    }

    private static Integer getExcelValueAnnotationValue(Field f) {
        return Integer.valueOf(f.getAnnotation(ExcelValue.class).value());
    }

}
