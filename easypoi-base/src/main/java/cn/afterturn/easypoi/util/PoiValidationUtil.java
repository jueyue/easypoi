/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package cn.afterturn.easypoi.util;

import cn.afterturn.easypoi.excel.annotation.Excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;

/**
 * HIBERNATE 校验工具类
 * @author JueYue
 *  2015年11月11日 下午10:04:07
 */
public class PoiValidationUtil {

    private final static Validator VALIDATOR;

    static {
        ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
        VALIDATOR = factory.getValidator();
    }

    public static String validation(Object obj, Class[] verfiyGroup) {
        Set<ConstraintViolation<Object>> set = null;
        if(verfiyGroup != null){
            set = VALIDATOR.validate(obj,verfiyGroup);
        }else{
            set = VALIDATOR.validate(obj);
        }
        if (set!= null && set.size() > 0) {
            return getValidateErrMsg(set);
        }
        return null;
    }

    private static String getValidateErrMsg(Set<ConstraintViolation<Object>> set) {
        StringBuilder builder = new StringBuilder();
        for (ConstraintViolation<Object> constraintViolation : set) {
            Class cls = constraintViolation.getRootBean().getClass();
            String fieldName = constraintViolation.getPropertyPath().toString();
            List<Field> fields = new ArrayList<>(Arrays.asList(cls.getDeclaredFields()));
            Class superClass = cls.getSuperclass();
            if (superClass != null) {
                fields.addAll(Arrays.asList(superClass.getDeclaredFields()));
            }
            for (Field field: fields) {
                if (field.getName().equals(fieldName) && field.isAnnotationPresent(Excel.class)) {
                    builder.append(field.getAnnotation(Excel.class).name());
                    break;
                }
            }
            builder.append(constraintViolation.getMessage()).append(",");
        }
        return builder.substring(0, builder.length() - 1);
    }

}
