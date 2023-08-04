原工具链接https://github.com/lowleveldesign/comon
由于霍尼plc还没修好本周开发hook辅助工具,磨刀不误砍柴工
自己改的windbg的com组件监视工具,可以监视所有类型com组件的创建行为,自动获取虚表函数地址,接续对应typelib获取函数名称,并生成人造函数地址符号,支持导出接口展示,此工具可以辅助hook开发.
修复了原工具断点卡死问题,添加自动索引typelib,根据clsid展示对应虚表符号功能
Com组件默认是没有调试符号,函数名称存在typelib中,需要动态解析生成,也就是说只有在com组件对象创建后才能找到对应接口的虚表.
函数地址通过这样计算得出 	ULONG64 pFunction = (ULONG64)(*(ULONG_PTR*)(vtbl + mod + fd->oVft));
虚表与接口的关系通过接口iid对应,clsid对应,同一对clsid和iid对应一个具体需要
启动方法.load dbgcom.dll;  !comon attach;g;
展示模块!cometa showm oleacc
索引模块!cometa index oleacc
展示clsid   !cometa showc {B5F8350B-0548-48B1-A6EE-88BD00B4A5E7}
