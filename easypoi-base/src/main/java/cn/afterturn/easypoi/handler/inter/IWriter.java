package cn.afterturn.easypoi.handler.inter;

import java.util.Collection;

/**
 * 大数据写出服务接口
 *
 * @author jueyue on 19-11-25.
 */
public interface IWriter<T> {
    /**
     * 获取输出对象
     *
     * @return
     */
    default public T get() {
        return null;
    }

    /**
     * 写入数据
     *
     * @param data
     * @return
     */
    public IWriter<T> write(Collection data);

    /**
     * 关闭流,完成业务
     *
     * @return
     */
    public T close();
}
