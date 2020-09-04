/**
 *  InsertCellDialog.js
 *
 *  @author yuanzhy
 *  @date   2020-09-04
 */
define([
    'common/main/lib/component/Window',
    'common/main/lib/component/ComboBox'
], function () { 'use strict';

    SSE.Views.InsertCellDialog = Common.UI.Window.extend(_.extend({
        options: {
            width: 214,
            height: 138,
            header: true,
            style: 'min-width: 214px;',
            cls: 'modal-dlg',
            buttons: ['ok', 'cancel']
        },

        initialize : function(options) {
            _.extend(this.options, options || {});

            this.template = [
                '<div class="box">',
                '<div id="insert-radio-shift-right" style="margin-bottom: 5px;"></div>',
                '<div id="insert-radio-shift-down" style="margin-bottom: 5px;"></div>',
                '<div id="insert-radio-row" style="margin-bottom: 5px;"></div>',
                '<div id="insert-radio-column"></div>',
                '</div>'
            ].join('');

            this.options.tpl = _.template(this.template)(this.options);

            Common.UI.Window.prototype.initialize.call(this, this.options);
        },

        render: function() {
            Common.UI.Window.prototype.render.call(this);

            this.isInsert = this.options.props === 'insert';

            this.radioShiftHorizontal = new Common.UI.RadioBox({
                el: $('#insert-radio-shift-right'),
                labelText: this.isInsert ? '向右移动单元格' : '向左移动单元格',
                name: 'asc-radio-insert-cells',
                // checked: this.options.props=='shiftRight'
            });

            this.radioShiftVertical = new Common.UI.RadioBox({
                el: $('#insert-radio-shift-down'),
                labelText: this.isInsert ? '向下移动单元格': '向上移动单元格',
                name: 'asc-radio-insert-cells',
                // checked: this.options.props=='shiftDown'
            });
            this.radioRow = new Common.UI.RadioBox({
                el: $('#insert-radio-row'),
                labelText: '整行',
                name: 'asc-radio-insert-cells',
                // checked: this.options.props=='row'
            });
            this.radioColumn = new Common.UI.RadioBox({
                el: $('#insert-radio-column'),
                labelText: '整列',
                name: 'asc-radio-insert-cells',
                // checked: this.options.props=='column'
            });

            this.isInsert ? this.radioShiftVertical.setValue(true) : this.radioShiftHorizontal.setValue(true);
            // (this.options.props=='rows') ? this.radioRows.setValue(true) : this.radioColumns.setValue(true);

            var $window = this.getChild();
            $window.find('.dlg-btn').on('click', _.bind(this.onBtnClick, this));

            var radios = [this.radioShiftHorizontal, this.radioShiftVertical, this.radioRow, this.radioColumn];

            this.onkeydown = function (e) {
                var i;
                if (e.keyCode === Common.UI.Keys.UP) {
                    for (i = 1; i < radios.length; i++) {
                        if (radios[i].getValue()) {
                            radios[i].setValue(false);
                            radios[i-1].setValue(true);
                            break;
                        }
                    }
                } else if (e.keyCode === Common.UI.Keys.DOWN) {
                    for (i = 0; i < radios.length - 1; i++) {
                        if (radios[i].getValue()) {
                            radios[i].setValue(false);
                            radios[i+1].setValue(true);
                            break;
                        }
                    }
                }
            }
            $(document).on('keydown',       this.onkeydown);
        },

        _handleInput: function(state) {
            if (this.options.handler) {
                this.options.handler.call(this, this, state);
            }
            $(document).off('keydown',       this.onkeydown);
            this.close();
        },

        onBtnClick: function(event) {
            this._handleInput(event.currentTarget.attributes['result'].value);
        },

        getSettings: function() {
            if (this.radioShiftHorizontal.getValue()) {
                return 1; // Asc.c_oAscInsertOptions.InsertCellsAndShiftRight or c_oAscDeleteOptions.DeleteCellsAndShiftLeft
            }
            if (this.radioShiftVertical.getValue()) {
                return 2; // Asc.c_oAscInsertOptions.InsertCellsAndShiftDown or c_oAscDeleteOptions.DeleteCellsAndShiftTop
            }
            if (this.radioRow.getValue()) {
                return 4; // Asc.c_oAscInsertOptions.InsertRows or c_oAscDeleteOptions.DeleteRows
            }
            if (this.radioColumn.getValue()) {
                return 3; // Asc.c_oAscInsertOptions.InsertColumns or c_oAscDeleteOptions.DeleteColumns
            }
        },

        onPrimary: function() {
            this._handleInput('ok');
            return false;
        },
    }, SSE.Views.InsertCellDialog || {}))
});