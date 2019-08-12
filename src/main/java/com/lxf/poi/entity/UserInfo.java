package com.lxf.poi.entity;

import java.io.Serializable;
import javax.persistence.*;
import lombok.Data;

@Data
@Table(name = "`user_info`")
public class UserInfo implements Serializable {
    @Column(name = "`name`")
    private String name;

    @Column(name = "`age`")
    private Integer age;

    @Column(name = "`address`")
    private String address;

    @Column(name = "`sort`")
    private Integer sort;

    private static final long serialVersionUID = 1L;
}