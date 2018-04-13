# Maven

## dependencymanagement和dependencys区别
　　dependencyManagement里只是声明依赖，并不实现引入，因此子项目需要显式的声明需要用的依赖。如果不在子项目中声明依赖，是不会从父项目中继承下来的；只有在子项目中写了该依赖项，并且没有指定具体版本，才会从父项目中继承该项，并且version和scope都读取自父pom;另外如果子项目中指定了版本号，那么会使用子项目中指定的jar版本。
　　dependencies即使在子模块中不写该依赖项，那么子模块仍然会从父项目中继承该依赖项（全部继承）。


## 相关资料
* [poi操作office](https://github.com/jiaowen194/Maven/blob/master/poi%E6%93%8D%E4%BD%9Coffice.md)
