#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安静模式启动器：重定向所有输出到日志文件，只显示最终结果
用法: python run_quiet.py --records-dir "..." --standard ISO20000 --audit-date 2026-02-01 --company-name "..."
"""

import os, sys, io, argparse
from datetime import datetime
from pathlib import Path

# 获取参数（不解析，直接传递给run_all.py）
import subprocess

def main():
    # 创建日志文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = f"output/log_{timestamp}.txt"
    os.makedirs("output", exist_ok=True)
    
    # 构建run_all.py命令
    run_all_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run_all.py")
    
    # 所有命令行参数原样传递给run_all.py
    cmd = [sys.executable, "-u", run_all_script] + sys.argv[1:]
    
    print(f"执行检查表生成...", flush=True)
    print(f"详细日志: {log_file}", flush=True)
    print(f"", flush=True)
    
    # 打开日志文件
    with open(log_file, 'w', encoding='utf-8') as log:
        log.write(f"ISO检查表生成日志 - {datetime.now()}\n")
        log.write(f"命令: {' '.join(cmd)}\n")
        log.write(f"{'='*80}\n\n")
        log.flush()
        
        # 执行run_all.py，实时输出到日志文件和控制台（可选）
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            encoding='utf-8',
            errors='replace'
        )
        
        final_output = []
        
        # 实时读取输出
        for line in process.stdout:
            # 写入日志文件
            log.write(line)
            log.flush()
            
            # 捕获最终输出行（包含"Output:"或".docx"的行）
            if "Output:" in line or ".docx" in line or "问题清单" in line:
                final_output.append(line.strip())
        
        process.wait()
    
    print(f"", flush=True)
    print(f"✅ 生成完成！", flush=True)
    print(f"", flush=True)
    print(f"📄 输出文件:", flush=True)
    for line in final_output:
        print(f"   {line}", flush=True)
    print(f"", flush=True)
    print(f"📋 详细日志: {log_file}", flush=True)
    print(f"", flush=True)

if __name__ == "__main__":
    main()
