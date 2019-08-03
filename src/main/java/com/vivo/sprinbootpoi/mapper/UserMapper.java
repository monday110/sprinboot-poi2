package com.vivo.sprinbootpoi.mapper;


import com.vivo.sprinbootpoi.entity.User;
import org.apache.ibatis.annotations.Select;
import java.util.List;

public interface UserMapper {
    @Select("select * from user")
    public List<User> getAll();

}
