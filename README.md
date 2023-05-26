# WpsFunction
用excel-dna 给wps 写自定义函数，环境用的是 excel-dna 1.7.0-rc版本。编译环境vs 2022 预览版。
里面有几个自己写的函数sthLink,sthReplace,sthFenJie。具体函数效果自己查看代码
感谢excelhome的wodewan大佬提供RegWps.cs代码，可以不需改动excel-dna 代码，直接给wps写自定义函数。
RegWps.cs代码的原理请去参考https://club.excelhome.net/thread-1628972-1-1.html
按F5 编译后在这个WpsFunction\bin\Debug\net6.0-windows\publish 目录下会有一个WpsFunction-AddIn-packed.xll 包，WPS直接加载这个包就可以用了。
注意，需要加载WpsFunction-AddIn-packed.xll 包后，关闭WPS，然后再打开才能正常使用自己的自定义函数。
再次感谢excelhome的wodewan大佬。
