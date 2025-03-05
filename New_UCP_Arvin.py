import tkinter
from tkinter import ttk
from for_data_02 import *
import json
import os
import sys
import subprocess
from tkinter import messagebox
import time

class UcpArvin(tkinter.Tk):
    def __init__(self):
        super().__init__()
        self.title("综合矿物资料管理系统")
        self.geometry("1100x780")

        # 配置路径（必须先于任何使用该属性的方法）
        self.json_dir = "for_json"
        os.makedirs(self.json_dir, exist_ok=True)

        # 初始化界面组件
        self._setup_ui()

        # 加载数据
        self.load_json_data()

        # 绑定事件
        self.treeview.bind("<<TreeviewSelect>>", self.on_tree_select)

    def _setup_ui(self):
        # 配置网格布局权重
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(3, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_columnconfigure(4, weight=0)
        self.grid_rowconfigure((0, 1, 2, 3), weight=1)

        # 创建侧边栏
        self.sidebar_frame = tkinter.Frame(self, bg="#2d2d2d", width=140)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")

        # 侧边栏组件
        self.logo_label = tkinter.Label(
            self.sidebar_frame,
            text="综合矿物资料管理系统",
            bg="#2d2d2d", fg="white",
            font=("Arial", 14, "bold")
        )
        self.logo_label.pack(pady=(20, 10), padx=20, anchor="w")

        botten_01 = tkinter.Button(
            self.sidebar_frame,
            text="导入excel表格",
            bg="#3d3d3d", fg="white",
            relief="flat",
            command=self.import_excel
        )
        botten_01.pack(pady=5, padx=20, fill="x")

        botten_02 = tkinter.Button(
            self.sidebar_frame,
            text="导出excel表格",
            bg="#3d3d3d", fg="white",
            relief="flat",
            command=self.export_to_excel
        )
        botten_02.pack(pady=5, padx=20, fill="x")

        # ========== 主内容区 ==========
        self.main_content_frame = tkinter.Frame(self)
        self.main_content_frame.grid(row=0, column=1, rowspan=3, sticky="nsew")

        # 设置7:3的行权重比例
        self.main_content_frame.grid_rowconfigure(0, weight=1)  # Treeview区域
        self.main_content_frame.grid_rowconfigure(1, weight=0)  # 文本框区域
        self.main_content_frame.grid_columnconfigure(0, weight=1)

        # Treeview容器（70%高度）
        treeview_container = tkinter.Frame(self.main_content_frame)
        treeview_container.grid(row=0, column=0, sticky="nsew")
        treeview_container.grid_columnconfigure(0, weight=7)
        treeview_container.grid_rowconfigure(0, weight=7)

        self.treeview = ttk.Treeview(
            treeview_container,
            columns=("物品编号", "持有人", "采集地"),
            show="headings"
        )
        self.treeview.grid(row=0, column=0, sticky="nsew")

        # Treeview滚动条
        tree_scroll = ttk.Scrollbar(
            treeview_container,
            orient="vertical",
            command=self.treeview.yview
        )
        self.treeview.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=0, column=1, sticky="ns")

        # 配置Treeview列
        for col in ["物品编号", "持有人", "采集地"]:
            self.treeview.heading(col, text=col)
            self.treeview.column(col, width=150, anchor="center")

        # 文本框容器（30%高度）
        text_container = tkinter.Frame(self.main_content_frame)
        text_container.grid(row=1, column=0, sticky="nsew")
        text_container.grid_columnconfigure(0, weight=1)
        text_container.grid_rowconfigure(0, weight=1)

        self.info_text = tkinter.Text(
            text_container,
            wrap=tkinter.WORD,
            font=("Arial", 10)
        )
        self.info_text.grid(row=0, column=0, sticky="nsew")

        # 文本框滚动条
        text_scroll = ttk.Scrollbar(
            text_container,
            orient="vertical",
            command=self.info_text.yview
        )
        self.info_text.configure(yscrollcommand=text_scroll.set)
        text_scroll.grid(row=0, column=1, sticky="ns")

        # 插入示例数据，后面需要做应用间的对接，接入数据库的数据
        self.insert_sample_data()

        # 主内容区 - 输入区
        self.tab_control = ttk.Notebook(self)
        self.tab_control.grid(row=0, column=3, rowspan=4, sticky="nsew")
        # 创建第一个标签页
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab1, text="矿物信息")

        # 在标签页内创建结构化输入区域
        self._create_entry_fields(self.tab1)

    def _create_entry_fields(self, parent):
        # 创建带标签的输入字段组
        a_field = "序号"
        b_field = "物品编号"
        c_field = "持有人"
        s_field = "采集地"
        d_field = "薄片描述"

        # 配置父容器的网格权重
        parent.grid_columnconfigure(1, weight=1)  # 让输入区域可以扩展
        parent.grid_rowconfigure(5, weight=1)  # 让文本区域行可以扩展

        # 使用 grid 布局对齐
        ttk.Label(parent, text=f"{a_field}:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_a = ttk.Entry(parent, width=25)
        self.entry_a.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(parent, text=f"{b_field}:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_b = ttk.Entry(parent, width=25)
        self.entry_b.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(parent, text=f"{c_field}:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.entry_c = ttk.Entry(parent, width=25)
        self.entry_c.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(parent, text=f"{s_field}:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.entry_s = ttk.Entry(parent, width=25)
        self.entry_s.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(parent, text=f"{d_field}:").grid(row=5, column=0, padx=5, pady=5, sticky="ne")
        self.entry_d = text_area = tkinter.Text(
            parent,
            wrap=tkinter.WORD,
            height=8,
            width=25,
            font=("Arial", 10)
        )
        self.entry_d.grid(row=5, column=1, padx=5, pady=5, sticky="nsew")

        # 添加滚动条
        scrollbar = ttk.Scrollbar(
            parent,
            orient=tkinter.VERTICAL,
            command=text_area.yview
        )
        text_area.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=5, column=2, sticky="ns")

    def insert_sample_data(self):

        # 底部输入框和按钮
        bottom_frame = ttk.Frame(self)
        bottom_frame.grid(row=3, column=1, columnspan=2, sticky="nsew", padx=20, pady=20)

        self.entry = ttk.Entry(bottom_frame)
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.main_btn = ttk.Button(bottom_frame, text="查询", command=self.search_item_by_number)
        self.main_btn.pack(side="right")

        # 功能按钮组
        self.right_frame = ttk.Frame(self)
        self.right_frame.grid(row=0, column=4, rowspan=4, sticky="nsew", padx=10)

        # 扩展右侧功能区
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(1, weight=1)  # 标签页区域可扩展

        # 添加竖直排列的三个按钮
        self._create_vertical_buttons(self.right_frame)

        # 添加_create_vertial_buttons方法

    def _create_vertical_buttons(self, parent):
        # 按钮容器（竖直排列）
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=0, column=0, sticky="nsew", pady=5)

        # 五个功能按钮
        buttons = [
            ("添加矿物", self.add_data),
            ("删除矿物", self.remove_data),
            ("修改信息", self.fixed_data),
            ("查看图片", self.check_photo),
            ("查看资料", self.check_for_information)
        ]

        for idx, (text, command) in enumerate(buttons):
            btn = ttk.Button(
                button_frame,
                text=text,
                style="TButton",
                command=command
            )
            btn.grid(row=idx, column=0, padx=5, pady=3, sticky="ew")

        # 配置按钮容器行权重
        button_frame.columnconfigure(0, weight=1)
        for i in range(len(buttons)):
            button_frame.rowconfigure(i, weight=0)

    def load_json_data(self):
        """加载所有JSON文件到Treeview"""
        print(f"当前工作目录: {os.getcwd()}")
        print(f"正在扫描目录: {os.path.abspath(self.json_dir)}")

        # 清空现有数据
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # 遍历JSON目录
        file_list = os.listdir(self.json_dir)
        print(f"原始文件列表: {file_list}")  # 显示原始文件名

        # 添加文件名过滤调试
        valid_files = [f for f in file_list if f.lower().endswith(".json")]
        print(f"有效JSON文件: {valid_files}")  # 显示实际匹配到的文件

        for filename in file_list:
            # 添加文件名格式验证
            if not filename.lower().endswith(".json"):
                print(f"跳过非JSON文件: {filename}")
                continue

            filepath = os.path.join(self.json_dir, filename)
            print(f"正在处理文件: {filepath}")

            try:
                # 添加文件内容验证
                with open(filepath, "r", encoding="utf-8") as f:
                    raw_content = f.read()
                    print(f"文件原始内容（前100字符）:\n{raw_content[:100]}")  # 打印文件开头内容

                    data = json.loads(raw_content)
                    print(f"解析后数据: {data}")

                    # 验证必要字段
                    required_fields = ["物品编号", "持有人", "采集地"]
                    for field in required_fields:
                        if field not in data:
                            print(f"字段缺失: {field}")
                            raise KeyError(f"{field} 字段不存在")

                    # 插入数据
                    self.treeview.insert(
                        "",
                        "end",
                        values=(
                            data["物品编号"],
                            data["持有人"],
                            data["采集地"]
                        ),
                        tags=(filepath,)
                    )
                    print(f"成功插入: {data['物品编号']}")

            except json.JSONDecodeError as e:
                print(f"JSON解析失败 [{filename}]: {str(e)}")
            except KeyError as e:
                print(f"数据字段缺失 [{filename}]: {str(e)}")
            except Exception as e:
                print(f"其他错误 [{filename}]: {str(e)}")
            finally:
                print("-" * 50)  # 添加分隔线

    def on_tree_select(self, event):
        """处理Treeview选中事件"""
        selected = self.treeview.selection()
        if selected:
            item = self.treeview.item(selected[0])
            filepath = item["tags"][0]  # 从tags获取文件路径

            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    # 更新右侧输入框
                    self.entry_a.delete(0, tkinter.END)
                    self.entry_a.insert(0, data.get("矿物序号", ""))

                    self.entry_b.delete(0, tkinter.END)
                    self.entry_b.insert(0, data.get("物品编号", ""))

                    self.entry_c.delete(0, tkinter.END)
                    self.entry_c.insert(0, data.get("持有人", ""))

                    self.entry_s.delete(0, tkinter.END)
                    self.entry_s.insert(0, data.get("采集地", ""))

                    self.entry_d.delete("1.0", tkinter.END)
                    self.entry_d.insert("1.0", data.get("薄片描述", ""))
                    self.update_info_text(data)

            except Exception as e:
                print(f"读取选中文件出错: {str(e)}")

    def update_info_text(self, data):
        """更新信息显示文本框"""
        # 清空原有内容
        self.info_text.config(state=tkinter.NORMAL)
        self.info_text.delete(1.0, tkinter.END)

        # 创建格式化文本
        info_str = ""
        info_str += "▌矿物详细信息\n\n"
        info_str += f"• 矿物序号：{data.get('矿物序号', '无')}\n"
        info_str += f"• 物品编号：{data.get('物品编号', '无')}\n"
        info_str += f"• 持有人　：{data.get('持有人', '无')}\n"
        info_str += f"• 采集地　：{data.get('采集地', '无')}\n"
        info_str += "\n▌薄片描述\n"
        info_str += "-" * 40 + "\n"
        info_str += data.get('薄片描述', '暂无描述')

        # 插入文本并设置为只读
        self.info_text.insert(tkinter.END, info_str)
        self.info_text.config(state=tkinter.DISABLED)

        # 添加标签配置（在初始化时添加）
        self.info_text.tag_config('header', font=('微软雅黑', 12, 'bold'), foreground='#2C3E50')
        self.info_text.tag_config('field', font=('宋体', 10))
        self.info_text.tag_config('divider', font=('宋体', 8), foreground='#95A5A6')

    def add_data(self):
        xuhao = self.entry_a.get()
        bianhao = self.entry_b.get()
        chiyouren = self.entry_c.get()
        caijidi = self.entry_s.get()
        miaoshu = self.entry_d.get("1.0", tkinter.END).strip()
        add_data_01(xuhao, bianhao, chiyouren, caijidi, miaoshu)
        # 清空所有输入组件
        self.entry_a.delete(0, tkinter.END)
        self.entry_b.delete(0, tkinter.END)
        self.entry_c.delete(0, tkinter.END)
        self.entry_s.delete(0, tkinter.END)
        self.entry_d.delete("1.0", tkinter.END)
        # 添加数据后刷新Treeview
        self.load_json_data()

    def remove_data(self):
        # 获取当前选中的文件路径
        selected = self.treeview.selection()
        if selected:
            item = self.treeview.item(selected[0])
            filepath = item["tags"][0]

            # 删除文件
            try:
                os.remove(filepath)
                self.load_json_data()  # 刷新数据
                # 清空输入框
                self.entry_a.delete(0, tkinter.END)
                self.entry_b.delete(0, tkinter.END)
                self.entry_c.delete(0, tkinter.END)
                self.entry_s.delete(0, tkinter.END)
                self.entry_d.delete("1.0", tkinter.END)
            except Exception as e:
                print(f"删除文件失败: {str(e)}")

    def fixed_data(self):
        # 获取当前选中的文件路径
        selected = self.treeview.selection()
        if selected:
            item = self.treeview.item(selected[0])
            old_filepath = item["tags"][0]

            # 获取新数据
            new_data = {
                "矿物序号": self.entry_a.get(),
                "物品编号": self.entry_b.get(),
                "持有人": self.entry_c.get(),
                "采集地": self.entry_s.get(),
                "薄片描述": self.entry_d.get("1.0", tkinter.END).strip()
            }

            # 删除旧文件，保存新文件
            try:
                os.remove(old_filepath)
                new_filename = f"{new_data['物品编号']}.json"
                new_filepath = os.path.join(self.json_dir, new_filename)
                with open(new_filepath, "w", encoding="utf-8") as f:
                    json.dump(new_data, f, ensure_ascii=False, indent=2)
                self.load_json_data()  # 刷新数据
            except Exception as e:
                print(f"修改数据失败: {str(e)}")

    def check_photo(self):
        # 获取物品编号
        item_number = self.entry_b.get().strip()

        if not item_number:
            messagebox.showwarning("提示", "请先选择或输入物品编号")
            return

        # 配置图片目录路径
        photo_dir = "for_photo"
        os.makedirs(photo_dir, exist_ok=True)  # 自动创建目录（如果不存在）

        # 构建目标文件夹路径
        target_folder = os.path.join(photo_dir, item_number)

        if os.path.isdir(target_folder):
            try:
                # 使用系统命令打开文件夹
                if sys.platform == "win32":
                    os.startfile(target_folder)
                elif sys.platform == "darwin":
                    subprocess.call(["open", target_folder])
                else:
                    subprocess.call(["xdg-open", target_folder])
            except Exception as e:
                messagebox.showerror("打开失败", f"无法打开文件夹:\n{str(e)}")
        else:
            messagebox.showinfo("未找到",
                                f"在 {photo_dir} 目录中\n未找到编号为 {item_number} 的文件夹",
                                detail=f"请确认已创建对应文件夹：\n{target_folder}")

    def check_for_information(self):
        # 获取物品编号
        item_number = self.entry_b.get().strip()

        if not item_number:
            messagebox.showwarning("提示", "请先选择或输入物品编号")
            return

        # 配置图片目录路径
        information_dir = "for_word_information"
        os.makedirs(information_dir, exist_ok=True)  # 自动创建目录（如果不存在）

        # 构建目标文件夹路径
        target_folder = os.path.join(information_dir, item_number)

        if os.path.isdir(target_folder):
            try:
                # 使用系统命令打开文件夹
                if sys.platform == "win32":
                    os.startfile(target_folder)
                elif sys.platform == "darwin":
                    subprocess.call(["open", target_folder])
                else:
                    subprocess.call(["xdg-open", target_folder])
            except Exception as e:
                messagebox.showerror("打开失败", f"无法打开文件夹:\n{str(e)}")
        else:
            messagebox.showinfo("未找到",
                                f"在 {information_dir} 目录中\n未找到编号为 {item_number} 的文件夹",
                                detail=f"请确认已创建对应文件夹：\n{target_folder}")

    def search_item_by_number(self):
        """根据物品编号搜索并选中对应的Treeview项"""
        search_number = self.entry.get().strip()  # 获取搜索框内容

        # 遍历所有Treeview项
        for item in self.treeview.get_children():
            values = self.treeview.item(item, "values")
            if len(values) >= 2 and values[1] == search_number:  # values[1]对应物品编号
                self.treeview.selection_set(item)  # 选中匹配项
                self.treeview.see(item)  # 滚动到可见位置
                return

        # 没有找到时显示提示
        messagebox.showinfo("提示", f"未找到编号为 {search_number} 的记录")

    def import_excel(self):
        """将Excel数据转换为JSON文件"""
        try:
            import pandas as pd

            # 配置文件路径
            excel_path = "information.xlsx"
            output_dir = "for_json"
            os.makedirs(output_dir, exist_ok=True)

            # 读取Excel数据
            df = pd.read_excel(excel_path)
            print(f"成功读取Excel文件，共{len(df)}条记录")

            # 验证必要列
            required_columns = ["矿物序号", "物品编号", "持有人", "采集地", "薄片描述"]
            if not all(col in df.columns for col in required_columns):
                missing = [col for col in required_columns if col not in df.columns]
                raise ValueError(f"Excel缺少必要列: {missing}")

            # 转换处理
            success_count = 0
            for index, row in df.iterrows():
                # 构建数据字典
                item_data = {
                    "矿物序号": str(row["矿物序号"]).strip(),
                    "物品编号": str(row["物品编号"]).strip(),
                    "持有人": str(row["持有人"]).strip(),
                    "采集地": str(row["采集地"]).strip(),
                    "薄片描述": str(row["薄片描述"]).strip()
                }

                # 验证关键字段
                if not item_data["物品编号"]:
                    print(f"第{index + 2}行物品编号为空，已跳过")
                    continue

                # 生成文件名
                filename = f"{item_data['物品编号']}.json"
                save_path = os.path.join(output_dir, filename)

                # 保存JSON文件
                with open(save_path, "w", encoding="utf-8") as f:
                    json.dump(item_data, f, ensure_ascii=False, indent=2)
                success_count += 1

            # 显示结果
            self.load_json_data()  # 刷新Treeview
            messagebox.showinfo(
                "导入完成",
                f"成功转换 {success_count}/{len(df)} 条数据\n"
                f"JSON文件保存至：{os.path.abspath(output_dir)}"
            )

        except FileNotFoundError:
            messagebox.showerror("文件未找到", f"无法找到Excel文件：{excel_path}")
        except Exception as e:
            messagebox.showerror("导入失败", f"发生错误：{str(e)}")

    def export_to_excel(self):
        """将JSON数据导出为Excel文件"""
        try:
            import pandas as pd

            # 配置路径
            json_dir = "for_json"
            excel_name = f"UCP生成事件_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            excel_path = os.path.join(os.getcwd(), excel_name)

            # 收集数据
            data_list = []
            required_fields = ["矿物序号", "物品编号", "持有人", "采集地", "薄片描述"]

            # 遍历JSON文件
            for filename in os.listdir(json_dir):
                if filename.lower().endswith(".json"):
                    filepath = os.path.join(json_dir, filename)
                    with open(filepath, "r", encoding="utf-8") as f:
                        try:
                            data = json.load(f)
                            # 验证字段完整性
                            if all(field in data for field in required_fields):
                                data_list.append({
                                    "矿物序号": data["矿物序号"],
                                    "物品编号": data["物品编号"],
                                    "持有人": data["持有人"],
                                    "采集地": data["采集地"],
                                    "薄片描述": data["薄片描述"]
                                })
                            else:
                                print(f"文件 {filename} 缺少必要字段，已跳过")
                        except json.JSONDecodeError:
                            print(f"文件 {filename} 格式错误，已跳过")

            # 生成DataFrame
            if len(data_list) == 0:
                messagebox.showwarning("无数据", "未找到有效JSON数据")
                return

            df = pd.DataFrame(data_list)

            # 按指定顺序排列列
            df = df[required_fields]

            # 写入Excel
            df.to_excel(
                excel_path,
                index=False,
                engine="openpyxl",
                sheet_name="矿物数据"
            )

            # 显示结果
            messagebox.showinfo(
                "导出成功",
                f"已成功导出 {len(data_list)} 条数据\n"
                f"文件保存至：{excel_path}"
            )

        except PermissionError:
            messagebox.showerror(
                "写入失败",
                "请关闭已打开的Excel文件后重试"
            )
        except Exception as e:
            messagebox.showerror(
                "导出失败",
                f"发生错误：{str(e)}"
            )


if __name__ == "__main__":
    app = UcpArvin()
    app.mainloop()