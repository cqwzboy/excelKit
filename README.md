# 介绍
Excel导入/导出通用工具

# 使用
## maven依赖
    <dependency>  
         <groupId>com.efficient</groupId>  
         <artifactId>excelKit</artifactId>  
         <version>1.0</version>  
     </dependency>
     
     仓库地址：http://47.106.201.17:8081/nexus/content/groups/public
     
     
## java代码
### 注解
* DateFormat
    
    format 时间格式化，如yyyy-MM-dd HH:mm:ss

* ExcelHeader

    title 标题
    
    order 排序 -128到127 值越小越靠前
    
    width 单元格宽度

* ExcelTypeHandler
    
    clazz 字段类型转换器Class，类型转换器需要实现接口 ExcelTypeHandler
    
        public interface ExcelTypeHandler<T> {
        
            /**
            * 导入时的类型转换
            */
            T onImport(String keyword);
        
            /**
            * 导出时的类型转换
            */
            String onExport(T t);
        
        }

### 工具类
ExcelUtil
    
    importData(String path, String sheetName, Class<T> clazz)
    
    exportData(String path, String sheetName, List<T> list)
    
## 示例
实体类：InputBean

    /**
     * Excel导入bean
     *
     * @author qinqin Fu
     * @since 2018-10-11
     * */
    @Data
    @ToString
    public class InputBean {
        /**
         * 分类
         * */
        @ExcelHeader(title = "二级分类", order = 0)
        private String category;
        /**
         * 关键词
         * */
        @ExcelHeader(title = "关键词", order = 1)
        private String keyword;
        /**
         * 平台
         * */
        @ExcelHeader(title = "平台", width = 10, order = 2)
        private String platform;
        /**
         * 作者
         * */
        @ExcelHeader(title = "作者", order = 3)
        private String author;
        /**
         * 上级作者
         * */
        @ExcelHeader(title = "上级作者", order = 4)
        private String superAuthor;
        /**
         * 态度 正面，负面
         * */
        @ExcelHeader(title = "正负面", width = 10, order = 5)
        private String attitude;
        /**
         * 内容
         * */
        @ExcelHeader(title = "内容", width = 100, order = 6)
        private String content;
        /**
         * 内容连接
         * */
        @ExcelHeader(title = "链接", order = 7)
        private String link;
        /**
         * 发布工具 例如 苹果7，小米max等
         * */
        @ExcelHeader(title = "发布工具", order = 8)
        private String publishTool;
        /**
         * 转发次数
         * */
        @ExcelHeader(title = "转发", width = 10, order = 9)
        private Integer forwardNum;
        /**
         * 评论次数
         * */
        @ExcelHeader(title = "评论", width = 10, order = 10)
        private Integer commentNum;
        /**
         * 发布时间
         * */
        @ExcelHeader(title = "发布时间", order = 11)
        @DateFormat(format = "yyyy/MM/dd")
        private Date publishTime;
        /**
         * 是否加V
         * */
        @ExcelHeader(title = "是否加V", order = 12)
        @ExcelTypeHandler(clazz = YesOrNoHandler.class)
        private YesOrNo hasV;
        /**
         * 粉丝数量
         * */
        @ExcelHeader(title = "粉丝数", order = 13)
        private Integer fansNum;
        /**
         * 关注数
         * */
        @ExcelHeader(title = "关注数", order = 14)
        private Integer followNum;
        /**
         * 微博数
         * */
        @ExcelHeader(title = "微博数", order = 15)
        private Integer weiboNum;
        /**
         * 点赞数
         * */
        @ExcelHeader(title = "点赞数", order = 16)
        private Integer likeNum;
        /**
         * 认证信息
         * */
        @ExcelHeader(title = "认证信息", order = 17)
        private String certInfo;
        /**
         * 性别
         * */
        @ExcelHeader(title = "性别", width = 10, order = 18)
        @ExcelTypeHandler(clazz = SexHandler.class)
        private Sex sex;
        /**
         * 省份
         * */
        @ExcelHeader(title = "省份", order = 19)
        private String province;
        /**
         * 城市
         * */
        @ExcelHeader(title = "城市", order = 20)
        private String city;
    }
    
调用类

    @Slf4j
    public class MainStart {
    
        public static void main(String[] args){
            log.info("== 开始导入 ==");
            String path = "F:\\工作\\2018\\10\\8 私活走起\\你好\\你好\\【原始数据】自如网_舆情动态_微博信息2018-9-29\\自如网_舆情动态_微博信息2018-9-29.xlsx";
            List<InputBean> list = ExcelUtil.importData(path, "敏感舆情", InputBean.class);
            for (InputBean inputBean : list) {
                log.info(inputBean.toString());
            }
    
            log.info("== 开始导出 ==");
            path = "F:\\工作\\2018\\10\\8 私活走起\\你好\\你好\\【原始数据】自如网_舆情动态_微博信息2018-9-29\\test.xls";
            ExcelUtil.exportData(path, "敏感舆情", list);
        }
    
    }
    
转换器

    /**
     * Sex枚举类转换器
     *
     * @author qinqin Fu
     * @since 2018-10-11
     */
    public class SexHandler implements ExcelTypeHandler<Sex> {
        @Override
        public Sex onImport(String keyword) {
            try{
                return Sex.keywordOf(keyword);
            }catch (Exception e){
                return null;
            }
        }
    
        @Override
        public String onExport(Sex sex) {
            return sex.keyword();
        }
    }
    
枚举类

    /**
     * 性别
     *
     * @author qinqin Fu
     * @since 2018-10-11
     */
    public enum Sex {
        UNKNOWN("未知"),
        MAN("男"),
        WUMAN("女"),
        ;
    
        private String keyword;
    
        Sex(String keyword){
            this.keyword = keyword;
        }
    
        public String keyword() {
            return this.keyword;
        }
    
        public static Sex keywordOf(String keyword) {
            if(StringUtils.isBlank(keyword)){
                return null;
            }
    
            for (Sex sex : Sex.values()) {
                if(keyword.equals(sex.keyword())){
                    return sex;
                }
            }
    
            return null;
        }
    }