package com.zqh.excel.linkage;

/**
 * Created by OrangeKiller on 2017/3/22.
 */
public class AreaInfo {

    private Long id;

    private Long parentId;

    private String name;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public Long getParentId() {
        return parentId;
    }

    public void setParentId(Long parentId) {
        this.parentId = parentId;
    }

    public String getName() {
        return name == null ? null : name.trim();
    }

    public void setName(String name) {
        this.name = name == null ? null : name.trim();
    }
}
