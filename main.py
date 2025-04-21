import traceback
import os
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import base64
from astrbot.api.star import Star, register, Context
from astrbot.api.event import filter, AstrMessageEvent
from astrbot.api.event.filter import PermissionType
from astrbot.api import logger
import astrbot.api.message_components as Comp

# 尝试导入特定平台事件类
try:
    from astrbot.core.platform.sources.aiocqhttp.aiocqhttp_message_event import AiocqhttpMessageEvent
except ImportError:
    logger.warning("无法导入 AiocqhttpMessageEvent，特定平台功能可能无法使用")
    AiocqhttpMessageEvent = None

@register(
    "astrbot_plugin_group_information",
    "Futureppo",
    "导出群成员信息为Excel表格",
    "1.0.5",
    "https://github.com/Futureppo/astrbot_plugin_group_information"
)
class GroupInformationPlugin(Star):
    def __init__(self, context: Context):
        super().__init__(context)
        # 设置插件临时目录
        self.temp_dir = Path(__file__).parent / "temp"
        try:
            self.temp_dir.mkdir(parents=True, exist_ok=True)
            logger.info(f"临时目录设置为：{self.temp_dir}（不推荐）")
        except Exception as e:
            logger.error(f"创建临时目录失败：{e}")
        logger.info("群信息插件加载完成")

    def _format_timestamp(self, timestamp):
        """格式化时间戳为可读时间"""
        if timestamp and isinstance(timestamp, (int, float)) and timestamp > 0:
            try:
                return datetime.fromtimestamp(float(timestamp)).strftime('%Y-%m-%d %H:%M:%S')
            except Exception as e:
                logger.warning(f"时间戳格式化失败：{e}")
        return None
        
    def _clean_excel_invalid_chars(self, text):
        """清理Excel不支持的特殊字符"""
        if not text or not isinstance(text, str):
            return text
            
        # 替换或移除Excel不支持的字符
        cleaned_text = ""
        for char in text:
            # 检查字符是否为控制字符（ASCII码小于32）或某些特殊字符
            if ord(char) < 32 or char in ['\x00', '\x01', '\x02', '\x03']:
                continue  # 跳过这些字符
            cleaned_text += char
        return cleaned_text

    @filter.command("导出群数据")
    async def export_group_data(self, event: AstrMessageEvent):
        """导出群成员信息到Excel文件"""
        # 检查是否在群聊中
        group_id_str = event.get_group_id()
        if not group_id_str:
            yield event.plain_result("请在群聊中使用此命令")
            return

        try:
            group_id = int(group_id_str)
        except ValueError:
            logger.warning(f"无效的群号格式：{group_id_str}")
            yield event.plain_result("无法识别的群号格式")
            return

        # 检查平台支持情况
        if event.get_platform_name() != "aiocqhttp":
            yield event.plain_result("此功能仅支持QQ平台")
            return

        if AiocqhttpMessageEvent is None or not isinstance(event, AiocqhttpMessageEvent):
            logger.error("事件类型不匹配或平台组件不可用")
            yield event.plain_result("内部错误：平台组件不可用")
            return

        client = event.bot
        if not client or not hasattr(client, 'api') or not hasattr(client.api, 'call_action'):
            logger.error("无法获取有效客户端")
            yield event.plain_result("无法获取平台连接，请检查日志")
            return

        # 获取群成员列表
        try:
            logger.info(f"获取群 {group_id} 的成员列表...")
            result = await client.api.call_action("get_group_member_list", group_id=group_id, no_cache=True)
            if isinstance(result, list):
                members = result
                logger.info(f"成功获取 {len(members)} 名成员")
            else:
                logger.warning(f"API返回格式异常：{type(result)}")
                yield event.plain_result("获取成员列表失败，请检查日志")
                return
        except Exception as e:
            logger.error(f"获取成员列表失败：{e}")
            yield event.plain_result("获取成员列表时出错")
            return

        if not members:
            yield event.plain_result("成员列表为空")
            return

        # 处理成员数据
        processed_members = []
        for member in members:
            if not isinstance(member, dict):
                logger.warning("发现非字典类型的成员数据，已跳过")
                continue

            member_copy = member.copy()
            # 清理可能包含特殊字符的字段
            for field in ["nickname", "card", "title"]:
                if field in member_copy and member_copy[field]:
                    member_copy[field] = self._clean_excel_invalid_chars(member_copy[field])
            
            member_copy["join_time"] = self._format_timestamp(member.get("join_time"))
            member_copy["last_sent_time"] = self._format_timestamp(member.get("last_sent_time"))
            member_copy["title_expire_time"] = self._format_timestamp(member.get("title_expire_time", 0))
            member_copy["shut_up_timestamp"] = self._format_timestamp(member.get("shut_up_timestamp", 0))
            processed_members.append(member_copy)

        if not processed_members:
            yield event.plain_result("无有效成员数据")
            return

        # 生成Excel文件到内存中
        try:
            df = pd.DataFrame(processed_members)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            file_name = f"群{group_id}成员列表_{ts}.xlsx"
            output_buffer = BytesIO()
            df.to_excel(output_buffer, index=False, engine='openpyxl', sheet_name=f"Group_{group_id}")
            output_buffer.seek(0)  # 重置缓冲区指针到开头
            file_content = output_buffer.getvalue()
            logger.info(f"Excel文件已生成到内存中：{file_name}")
        except Exception as e:
            logger.error(f"生成Excel失败：{e}\n{traceback.format_exc()}")
            yield event.plain_result("生成Excel时出错，请检查日志")
            return

        # 发送文件 - 方法1：原有发送方案
        try:
            logger.info("方法1：尝试通过原有方案发送文件 (export_group_data)...")
            file_component = Comp.File(
                name=f"群{group_id}成员列表_{ts}.xlsx",
                file=BytesIO(file_content)
            )
            yield event.chain_result([file_component])
            logger.info(f"方法1：文件已发送至群 {group_id}")
            return  # 如果成功，直接返回
        except Exception as e:
            logger.error(f"方法1：发送文件失败 (export_group_data)：{e}")

        # 发送文件 - 方法2：备用方案，只有在方法1失败时执行，仅上传文件不做任何提示
        try:
            logger.info("方法2：尝试使用适配器接口 upload_group_file 上传文件 (export_group_data)...")
            # 将字节流编码为 Base64 字符串
            file_content_base64 = base64.b64encode(file_content).decode('utf-8')
            # 构造符合 NapCat/go-cqhttp 格式的文件参数
            upload_result = await client.api.call_action(
                "upload_group_file",
                group_id=group_id,
                file=f"base64://{file_content_base64}",  # 使用 base64:// 前缀，符合 go-cqhttp 格式
                name=file_name
            )
            logger.info(f"方法2：upload_group_file 返回值 (export_group_data)：{upload_result}")
            logger.info(f"方法2：文件上传操作完成 (export_group_data)：{file_name}")
            # 发送成功导出的提示信息
            yield event.plain_result(f"已成功导出群 {group_id} 的 {len(processed_members)} 名成员信息")
            logger.info(f"方法2：已发送成功导出提示信息：群 {group_id} 的 {len(processed_members)} 名成员")
        except Exception as upload_e:
            logger.error(f"方法2：文件上传失败 (export_group_data)：{upload_e}")

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("导出所有群数据")
    async def export_all_groups_data(self, event: AstrMessageEvent):
        """导出所有群的成员信息到Excel文件"""
        # 检查平台支持情况
        if event.get_platform_name() != "aiocqhttp":
            yield event.plain_result("此功能仅支持QQ平台")
            return

        if AiocqhttpMessageEvent is None or not isinstance(event, AiocqhttpMessageEvent):
            logger.error("事件类型不匹配或平台组件不可用")
            yield event.plain_result("内部错误：平台组件不可用")
            return

        client = event.bot
        if not client or not hasattr(client, 'api') or not hasattr(client.api, 'call_action'):
            logger.error("无法获取有效客户端")
            yield event.plain_result("无法获取平台连接，请检查日志")
            return

        try:
            # 获取群列表
            logger.info("获取群列表...")
            group_list = await client.api.call_action("get_group_list")
            if not isinstance(group_list, list):
                logger.warning(f"API返回格式异常：{type(group_list)}")
                yield event.plain_result("获取群列表失败，请检查日志")
                return
            
            if not group_list:
                yield event.plain_result("未加入任何群聊")
                return
                
            # 创建Excel工作簿到内存中
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            file_name = f"all_groups_members_{ts}.xlsx"
            output_buffer = BytesIO()
            total_members = 0
            processed_groups = 0
            
            # 使用ExcelWriter创建多表格Excel文件
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                # 显示进度信息
                yield event.plain_result(f"开始导出 {len(group_list)} 个群的成员信息，请稍候...")
                
                # 遍历每个群
                for group in group_list:
                    group_id = group['group_id']
                    group_name = group['group_name']
                    
                    try:
                        # 获取群成员列表
                        logger.info(f"获取群 {group_id}({group_name}) 的成员列表...")
                        result = await client.api.call_action("get_group_member_list", group_id=group_id, no_cache=True)
                        
                        if not isinstance(result, list):
                            logger.warning(f"群 {group_id} API返回格式异常：{type(result)}")
                            continue
                            
                        members = result
                        if not members:
                            logger.info(f"群 {group_id} 成员列表为空")
                            continue
                            
                        # 处理成员数据
                        processed_members = []
                        for member in members:
                            if not isinstance(member, dict):
                                continue
                                
                            member_copy = member.copy()
                            # 清理可能包含特殊字符的字段
                            for field in ["nickname", "card", "title"]:
                                if field in member_copy and member_copy[field]:
                                    member_copy[field] = self._clean_excel_invalid_chars(member_copy[field])
                            
                            # 添加群信息
                            member_copy["group_id"] = group_id
                            member_copy["group_name"] = self._clean_excel_invalid_chars(group_name)
                            
                            # 格式化时间戳
                            member_copy["join_time"] = self._format_timestamp(member.get("join_time"))
                            member_copy["last_sent_time"] = self._format_timestamp(member.get("last_sent_time"))
                            member_copy["title_expire_time"] = self._format_timestamp(member.get("title_expire_time", 0))
                            member_copy["shut_up_timestamp"] = self._format_timestamp(member.get("shut_up_timestamp", 0))
                            
                            processed_members.append(member_copy)
                        
                        if processed_members:
                            # 创建DataFrame并写入Excel
                            df = pd.DataFrame(processed_members)
                            # 使用群号和群名作为表名，避免特殊字符问题
                            sheet_name = f"G{group_id}"
                            if len(sheet_name) > 31:  # Excel表名最长31字符
                                sheet_name = sheet_name[:31]
                            df.to_excel(writer, index=False, sheet_name=sheet_name)
                            
                            total_members += len(processed_members)
                            processed_groups += 1
                            logger.info(f"已导出群 {group_id}({group_name}) 的 {len(processed_members)} 名成员")
                    except Exception as e:
                        logger.error(f"处理群 {group_id} 时出错: {e}")
                        continue
            
            if processed_groups == 0:
                yield event.plain_result("未能成功导出任何群的成员信息")
                return
                
            output_buffer.seek(0)  # 重置缓冲区指针到开头
            file_content = output_buffer.getvalue()
            logger.info(f"Excel文件已生成到内存中：{file_name}")

            # 发送文件 - 方法1：原有发送方案
            try:
                logger.info("方法1：尝试通过原有方案发送文件 (export_all_groups_data)...")
                file_component = Comp.File(
                    name=f"所有群成员列表_{ts}.xlsx",
                    file=BytesIO(file_content)
                )
                yield event.chain_result([file_component])
                logger.info(f"方法1：已导出 {processed_groups} 个群，共 {total_members} 名成员的信息")
                yield event.plain_result(f"已成功导出 {processed_groups} 个群，共 {total_members} 名成员的信息")
                return  # 如果成功，直接返回
            except Exception as e:
                logger.error(f"方法1：发送文件失败 (export_all_groups_data)：{e}")

            # 发送文件 - 方法2：备用方案，只有在方法1失败时执行，仅上传文件不做任何提示
            try:
                group_id_str = event.get_group_id()
                if not group_id_str:
                    logger.error("方法2：无法获取群ID，无法上传文件 (export_all_groups_data)")
                else:
                    try:
                        group_id = int(group_id_str)
                        logger.info("方法2：尝试使用适配器接口 upload_group_file 上传文件 (export_all_groups_data)...")
                        # 将字节流编码为 Base64 字符串
                        file_content_base64 = base64.b64encode(file_content).decode('utf-8')
                        # 构造符合 NapCat/go-cqhttp 格式的文件参数
                        upload_result = await client.api.call_action(
                            "upload_group_file",
                            group_id=group_id,
                            file=f"base64://{file_content_base64}",  # 使用 base64:// 前缀，符合 go-cqhttp 格式
                            name=file_name
                        )
                        logger.info(f"方法2：upload_group_file 返回值 (export_all_groups_data)：{upload_result}")
                        logger.info(f"方法2：文件上传操作完成 (export_all_groups_data)：{file_name}")
                        # 发送成功导出的提示信息
                        yield event.plain_result(f"已成功导出 {processed_groups} 个群，共 {total_members} 名成员的信息")
                        logger.info(f"方法2：已发送成功导出提示信息：{processed_groups}个群，{total_members}名成员")
                    except ValueError:
                        logger.error(f"方法2：无效的群号格式 (export_all_groups_data)：{group_id_str}")
            except Exception as upload_e:
                logger.error(f"方法2：文件上传失败 (export_all_groups_data)：{upload_e}")
                
        except Exception as e:
            logger.error(f"导出所有群数据失败：{e}\n{traceback.format_exc()}")
            yield event.plain_result("导出所有群数据时出错，请检查日志")
            
    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("查看群列表")
    async def show_groups_info(self, event: AstrMessageEvent):
        """查看加入的所有群聊信息"""
        # 检查平台支持情况
        if event.get_platform_name() != "aiocqhttp":
            yield event.plain_result("此功能仅支持QQ平台")
            return

        if AiocqhttpMessageEvent is None or not isinstance(event, AiocqhttpMessageEvent):
            logger.error("事件类型不匹配或平台组件不可用")
            yield event.plain_result("内部错误：平台组件不可用")
            return

        client = event.bot
        if not client or not hasattr(client, 'api') or not hasattr(client.api, 'call_action'):
            logger.error("无法获取有效客户端")
            yield event.plain_result("无法获取平台连接，请检查日志")
            return

        try:
            logger.info("获取群列表...")
            group_list = await client.api.call_action("get_group_list")
            if not isinstance(group_list, list):
                logger.warning(f"API返回格式异常：{type(group_list)}")
                yield event.plain_result("获取群列表失败，请检查日志")
                return

            group_info = "\n".join(
                f"{g['group_id']}: {g['group_name']}" for g in group_list
            )
            info = f"【群列表】共加入{len(group_list)}个群：\n{group_info}"
            yield event.plain_result(info)
            logger.info(f"已发送群列表信息，共{len(group_list)}个群")
        except Exception as e:
            logger.error(f"获取群列表失败：{e}")
            yield event.plain_result("获取群列表时出错，请检查日志")

    async def terminate(self):
        """插件终止时清理临时文件"""
        logger.info("插件终止，开始清理临时文件...")
        try:
            if hasattr(self, 'temp_dir') and self.temp_dir.exists():
                for file in self.temp_dir.iterdir():
                    if file.is_file() and file.name.startswith("group_") and file.suffix == ".xlsx":
                        file.unlink()
                        logger.info(f"已删除：{file}")
        except Exception as e:
            logger.error(f"清理临时文件失败：{e}")
