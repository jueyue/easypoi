package cn.afterturn.easypoi.csv.handler;

/**
 * 保存数据
 * 鉴于CSV可能都是大数据,还是调用接口直接保存,避免内存占用
 *
 * @author by jueyue on 18-10-3.
 */
public interface ICsvSaveDataHandler<T> {

    /**
     * 保存数据
     *
     * @param t
     */
    public void save(T t);
}
