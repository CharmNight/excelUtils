# excelUtils
基于poi 练手的一个excel导出项目

## 新增
> 导出本地, 支持 同时写入多个sheet页,
>
> 如果 sheet 达到最大值 默认65535 可以自动扩展页 自动新增sheet名称规则: sheet_1\2\3
> 
>使用方式 详情见 test下 `AppendExcelBeanTest`

## 普通导出  详情见 test下 `UserInfoExcelBeanTest`

### 继承 `ExcelBean<UserInfoExcelBean>` 
```java
/**
 * @Author: CharmNight
 * @Date: 2020/8/19 0:03
 */
@ExcelName(value="用户信息表")
@Component
public class UserInfoExcelBean extends ExcelBean<UserInfoExcelBean> {
    @ExcelTitle("用户名")
    private String name;
    @ExcelTitle("状态")
    private String type;
    @ExcelTitle("创建时间")
    private String updatedAt;
    @ExcelTitle("创建人")
    private String createUser;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getUpdatedAt() {
        return updatedAt;
    }

    public void setUpdatedAt(String updatedAt) {
        this.updatedAt = updatedAt;
    }

    public String getCreateUser() {
        return createUser;
    }

    public void setCreateUser(String createUser) {
        this.createUser = createUser;
    }
}
```

ExcelName -> excel 中 sheetName
ExcelTitle -> excel 中标题

### 导出
```java
@Autowired
private UserInfoExcelBean bean;
// 导出本地
bean.exportToExcel("导出路径", 这里是个List<内容>);

// 如果需要通过http请求导出 参考 ReturnClient类

```

## 多sheet导出 详情见 test下 `AppendExcelBeanTest`
```java
// Map <String: sheetName, List<T>: 列的值>
Map<String, List<UserInfoExcelBean>> map = new HashMap<>();

bean.appendExportToExcel("导出路径", map);
```


## 代码阅读
utils 包下

- ExcelUtils      excel导出相关内容   (主要)
- ReturnClient    返回相关内容        (主要)
- GenericsUtils   反射获取父类泛型相关内容 (未使用)
- ReflectionUtils 反射获取get方法 (为了偷懒, 简单的拖了个类出来)

fields 包      -- 注解类
- ExcelName
- ExcelTitle 

beans 包
- ExcelBean -- 主要根基类 使用时需要继承的类 (我...取名取错了)
- UserInfoExcelBean -- 这个为了测试演示 
- ExcelExportEntity --  excel映射关系
