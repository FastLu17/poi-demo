package com.lxf.poi.mapper;

import com.lxf.poi.entity.UserInfo;
import tk.mybatis.mapper.common.Mapper;

import java.util.List;
import java.util.Map;

public interface UserInfoMapper extends Mapper<UserInfo> {
    List<Map<String, Object>> selectAllResultMap();

    Integer getSort();
}