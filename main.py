import traceback
import os
import pandas as pd
from pathlib import Path
from datetime import datetime
from astrbot.api.star import Star, register, Context
from astrbot.api.event import filter, AstrMessageEvent
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
    "1.0.3",
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

        # 生成Excel文件
        try:
            df = pd.DataFrame(processed_members)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            file_name = f"group_{group_id}_members_{ts}.xlsx"
            output_path = self.temp_dir / file_name
            # 使用群号作为文件名，避免特殊字符问题
            df.to_excel(output_path, index=False, engine='openpyxl', sheet_name=f"Group_{group_id}")
            logger.info(f"Excel文件已生成：{output_path}")
        except Exception as e:
            logger.error(f"生成Excel失败：{e}\n{traceback.format_exc()}")
            yield event.plain_result("生成Excel时出错，请检查日志")
            return

        # 发送文件
        try:
            file_component = Comp.File(
                name=f"群{group_id}成员列表_{ts}.xlsx",
                file=str(output_path)
            )
            yield event.chain_result([file_component])
            logger.info(f"文件已发送至群 {group_id}")
        except Exception as e:
            logger.error(f"发送文件失败：{e}")
            yield event.plain_result("发送文件时出错，请检查日志")

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
