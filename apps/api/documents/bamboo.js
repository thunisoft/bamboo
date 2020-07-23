/**
 * BambooOffice word api
 *
 * @author   BambooOffice Team
 * @version  1.0.0-beta.0
 * @date     2020-07-22
 */
;(function(window, document) {

    var frameName = 'frameEditor';
    // 菜单相关操作
    var menuMap = { // default is target@eventName
        // ============================ header ============================
        save: 'DE.controllers.Viewport.header.btnSave@click', // 保存
        print: 'DE.controllers.Viewport.header.btnPrint@click', // 打印
        undo: 'DE.controllers.Viewport.header.btnUndo@click', // 撤销
        redo: 'DE.controllers.Viewport.header.btnRedo@click', // 重做
        // fitPage: 'DE.controllers.Viewport.header.mnuitemFitPage@click', // 适合页面
        // fitWidth: 'DE.controllers.Viewport.header.mnuitemFitWidth@click', // 适合宽度

        // ============================ 主页 ============================
        copy: 'DE.controllers.Toolbar.toolbar.btnCopy@click', // 复制
        paste: 'DE.controllers.Toolbar.toolbar.btnPaste@click', // 粘贴

        incFontSize: 'DE.controllers.Toolbar.toolbar.btnIncFontSize@click', // 增大字号
        decFontSize: 'DE.controllers.Toolbar.toolbar.btnDecFontSize@click', // 减小字号
        bold: 'DE.controllers.Toolbar.toolbar.btnBold@click', // 粗体
        italic: 'DE.controllers.Toolbar.toolbar.btnItalic@click', // 倾斜
        underline: 'DE.controllers.Toolbar.toolbar.btnUnderline@click', // 下划线
        strikeout: 'DE.controllers.Toolbar.toolbar.btnStrikeout@click', // 删除线
        btnSuperscript: 'DE.controllers.Toolbar.toolbar.btnSuperscript@click', // 上标
        btnSubscript: 'DE.controllers.Toolbar.toolbar.btnSubscript@click',  // 下标

        incLeftOffset: 'DE.controllers.Toolbar.toolbar.btnIncLeftOffset@click', // 增加缩进
        decLeftOffset: 'DE.controllers.Toolbar.toolbar.btnDecLeftOffset@click', // 增加缩进
        lineSpace: {
            selectors: '#id-toolbar-btn-linespace [type=menuitem]',
            index: '#argumentsConfig.0',
            eventName: 'click',
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: Number, default: 1.15, items: [1.0, 1.15, 1.5, 2.0, 2.5, 3.0], description: '段线间距'}
            ]
        },
        alignLeft: 'DE.controllers.Toolbar.toolbar.btnAlignLeft@click', // 左对齐
        alignCenter: 'DE.controllers.Toolbar.toolbar.btnAlignCenter@click', // 居中对齐
        alignRight: 'DE.controllers.Toolbar.toolbar.btnAlignRight@click', // 右对齐
        alignJust: 'DE.controllers.Toolbar.toolbar.btnAlignJust@click', // 正当, 两端对齐

        clearStyle: 'DE.controllers.Toolbar.toolbar.btnClearStyle@click', // 清除样式

        // ============================ 插入 ============================
        insertBlankPage: 'DE.controllers.Toolbar.toolbar.btnBlankPage@click',
        // insertTable: '', // TODO

        // ============================ 布局 ============================
        // pageMargins: 'DE.controllers.Toolbar.toolbar.btnPageMargins.menu', //
        pageOrient: { // 布局: 竖向, 横向
            selectors: '#tlbtn-pageorient [type=menuitem]',
            index: '#argumentsConfig.0',
            eventName: 'click',
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: Boolean, default: true, items: [true, false], description: 'true:竖向,false:横向'}
            ]
        }
        // btnPageSize: 'DE.controllers.Toolbar.toolbar.btnPageMargins.menu', //

        // ============================ 参考 ============================
        // ============================ 协作 ============================
    }

    function BambooAPI() {}

    BambooAPI.create = function() {
        var iframes = document.getElementsByName(frameName);
        if (!iframes || iframes.length === 0) {
            throw new Error('BambooAPI对象必须在onAppReady中创建');
        }
        try {
            iframes[0].contentWindow.document;
            return new SameOriginBambooAPI(iframes[0]);
        } catch (e) {
            return new CrossOriginBambooAPI(iframes[0]);
        }
    }

    var _isString = function(arg) {
        return Object.prototype.toString.call(arg) === '[object String]';
    }
    var _getPropByPath = function(obj, path) {
        var tempObj = obj;
        path = path.replace(/\[(\w+)\]/g, '.$1');
        path = path.replace(/^\./, '');

        var keyArr = path.split('.');
        var i = 0;

        for (var len = keyArr.length; i < len - 1; ++i) {
            var key = keyArr[i];
            if (key in tempObj) {
                tempObj = tempObj[key];
            } else {
                throw new Error('[iView warn]: please transfer a valid prop path to form item!');
            }
        }
        return tempObj[keyArr[i]];
    }
    var _deepCopy = function (source) {
        var t = Object.prototype.toString.call(source)
        var o
        if (t === '[object Array]') {
            o = []
        } else if (t === '[object Object]') {
            o = {}
        } else {
            return source
        }
        if (t === '[object Array]') {
            for (var i = 0; i < source.length; i++) {
                o.push(_deepCopy(source[i]))
            }
        } else if (t === '[object Object]') {
            for (var key in source) {
                o[key] = _deepCopy(source[key])
            }
        }
        return o
    }

    var _translateIndex = function(eventObject, index, callArguments) {
        if (index === undefined || !_isString(index)) {
            return
        }
        if (index.indexOf('#argumentsConfig.') === 0) {
            var argConfigIndex = ~~index.substring('#argumentsConfig.'.length);
            var argConfig = eventObject.argumentsConfig[argConfigIndex];
            var callArg = callArguments[argConfigIndex] === undefined ? argConfig.default : argConfig.type(callArguments[argConfigIndex]);
            var itemIndex = argConfig.items.indexOf(callArg);
            if (itemIndex === -1) {
                throw new Error('Illegal argument, index: ' + argConfigIndex + ', value: ' + callArguments[argConfigIndex] + '. correct is ' + JSON.stringify(argConfig.items));
            }
            return itemIndex;
        }
    }

    for (var key in menuMap) {
        (function(fnName) {
            BambooAPI.prototype[fnName] = function () {
                var eventObject = menuMap[fnName];
                if (_isString(eventObject)) {
                    eventObject = {
                        target: eventObject.substring(0, eventObject.indexOf('@')),
                        eventName: eventObject.substring(eventObject.indexOf('@') + 1),
                        // eventParameters: [
                        //   {value: '#target', description: 'menu'},
                        //   {value: '#target.items', index: '#argumentsConfig.0', description: 'item' },
                        //   {value: true, description: 'state'}
                        // ],
                        // argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                        //   {type: Boolean, default: true, items: [true, false], description: 'true:竖向,false:横向'}
                        // ]
                    };
                } else {
                    eventObject = _deepCopy(eventObject);
                }

                var targetIndex = _translateIndex(eventObject, eventObject.index, arguments);
                if (targetIndex !== undefined) {
                    eventObject.index = targetIndex;
                }
                if (eventObject.eventParameters) {
                    for (var i = 0; i < eventObject.eventParameters.length; i++) {
                        var eventParam = eventObject.eventParameters[i];
                        var itemIndex = _translateIndex(eventObject, eventParam.index, arguments);
                        if (itemIndex !== undefined) {
                            eventParam.index = itemIndex;
                        }
                    }
                }
                this.trigger(eventObject);
            }
        })(key);
    }

    window.BambooAPI = BambooAPI;
    /**
     * 同源对象
     * @param frame iframe对象
     * @constructor
     */
    function SameOriginBambooAPI (frame) {
        this.frame = frame;
        this.frameWindow = frame.contentWindow;
        this.frameDocument = frame.contentDocument;
    }

    SameOriginBambooAPI.prototype = new BambooAPI();
    SameOriginBambooAPI.constructor = SameOriginBambooAPI;

    SameOriginBambooAPI.prototype.trigger = function(eventObject) {
        if (eventObject.selectors) {
            if (eventObject.index) {
                var target = this.frameDocument.querySelectorAll(eventObject.selectors);
                if (!target || target.length < eventObject.index) {
                    // 不做操作
                    throw new Error(eventObject.selectors + '['+eventObject.index+'] not found');
                }
                target[eventObject.index][eventObject.eventName]();
            } else {
                var target = this.frameDocument.querySelector(eventObject.selectors);
                if (!target) {
                    // 不做操作
                    throw new Error(eventObject.selectors + ' not found');
                }
                target[eventObject.eventName]();
            }
            return;
        }
        //
        var target = _getPropByPath(this.frameWindow, eventObject.target);
        if (!target) {
            throw new Error(eventObject.target + ' not found');
        }
        if (target.disabled) {
            return;
        }
        target.doToggle && target.doToggle();
        // if (target.options.hint) {
        //     var tip = target.btnEl.data('bs.tooltip');
        //     if (tip) {
        //         if (tip.dontShow===undefined)
        //             tip.dontShow = true;
        //         tip.hide();
        //     }
        // }
        eventObject.target = target;
        var parameters = [eventObject.eventName];
        if (eventObject.eventParameters) {
            for (var i = 0; i < eventObject.eventParameters.length; i++) {
                var eventParam = eventObject.eventParameters[i];
                if (_isString(eventParam.value) && eventParam.value.indexOf('#') === 0) {
                    var ref = eventParam.value.substring(1);
                    var arg = _getPropByPath(eventObject, ref);
                    if (!arg) {
                        throw new Error(eventParam.value + ' not found');
                    }
                    if (eventParam.index !== undefined) {
                        arg = arg[eventParam.index];
                    }
                    parameters.push(arg);
                } else {
                    parameters.push(eventParam.value);
                }
            }
        } else if (eventObject.index === undefined) {
            parameters.push(target);
        } else {
            target = target[eventObject.index];
            parameters.push(target);
        }
        target.trigger.apply(target, parameters);
    }
    /**
     * 跨域对象
     * @param frame iframe对象
     * @constructor
     */
    function CrossOriginBambooAPI (frame) {
        this.frame = frame;
        this.frameWindow = frame.contentWindow;
        this.frameDocument = frame.contentDocument;
    }

    CrossOriginBambooAPI.prototype = new BambooAPI();
    CrossOriginBambooAPI.constructor = CrossOriginBambooAPI;

    CrossOriginBambooAPI.prototype.trigger = function(eventObject) {
        var message = {
            source: 'bamboo',
            data: eventObject
        };
        // 为了兼容旧版IE, 先序列化为字符串吧
        this.frameWindow.postMessage(JSON.stringify(message), '*');
    }
})(window, document);