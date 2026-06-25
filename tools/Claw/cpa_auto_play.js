#!/usr/bin/env node
/**
 * CPA继续教育自动播放脚本
 * 自动化播放北京市注协CPA继续教育课程
 */

const { spawn, exec } = require('child_process');
const path = require('path');

// 配置
const CONFIG = {
  url: 'https://www.bicpaedu.com/cpajxjy/',
  sessionName: 'bicpaedu',
  screenshotDir: 'c:\\Users\\12844\\WorkBuddy\\Claw\\screenshots',
  maxLessonsPerRun: 3, // 每次运行最多播放3节课，避免超时
  waitTimeBetweenActions: 2000, // 操作间隔(ms)
  videoWaitTime: 30000, // 视频播放等待时间(ms)
};

// 辅助函数：执行agent-browser命令
function runAgentBrowserCmd(cmd, args = []) {
  return new Promise((resolve, reject) => {
    const command = `agent-browser ${cmd} ${args.join(' ')}`;
    exec(command, { maxBuffer: 1024 * 1024 }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(`命令执行失败: ${command}\n错误: ${error.message}\n输出: ${stdout}\n错误流: ${stderr}`));
        return;
      }
      resolve(stdout);
    });
  });
}

// 辅助函数：带超时的等待
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 截图并保存
async function takeScreenshot(filename) {
  try {
    const screenshotPath = path.join(CONFIG.screenshotDir, filename);
    await runAgentBrowserCmd('screenshot', [screenshotPath]);
    console.log(`✅ 截图已保存: ${screenshotPath}`);
    return screenshotPath;
  } catch (error) {
    console.log(`⚠️ 截图失败: ${error.message}`);
    return null;
  }
}

// 主要流程
async function main() {
  console.log('🚀 开始CPA继续教育自动播放任务...\n');
  const startTime = new Date();
  const results = {
    coursesPlayed: 0,
    errors: [],
    screenshots: [],
  };

  try {
    // 1. 创建截图目录
    const fs = require('fs');
    if (!fs.existsSync(CONFIG.screenshotDir)) {
      fs.mkdirSync(CONFIG.screenshotDir, { recursive: true });
    }

    // 2. 打开浏览器并导航到网站
    console.log('📂 打开浏览器...');
    await runAgentBrowserCmd('launch');
    await sleep(CONFIG.waitTimeBetweenActions);

    // 使用bicpaedu会话名来保存登录状态
    console.log(`🔗 导航到 ${CONFIG.url}`);
    await runAgentBrowserCmd('open', [CONFIG.url, '--session-name', CONFIG.sessionName]);
    await sleep(3000);

    // 3. 截图记录当前状态
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    await takeScreenshot(`bicpaedu_start_${timestamp}.png`);
    results.screenshots.push(`bicpaedu_start_${timestamp}.png`);

    // 4. 获取页面快照，查看结构
    console.log('📋 获取页面内容...');
    const snapshot = await runAgentBrowserCmd('snapshot', ['-i']);
    console.log('页面快照:', snapshot.substring(0, 500), '...\n');

    // 5. 查找"我的课堂"相关链接
    // 注意：根据实际页面结构调整选择器
    const myClassElements = [];
    
    // 尝试多种选择器
    const selectors = [
      'text=我的课堂',
      'text=我的课程',
      'text=继续学习',
      'text=听课',
      'a:has-text("我的课堂")',
      'a:has-text("继续学习")',
    ];

    for (const selector of selectors) {
      try {
        await sleep(500);
        const result = await runAgentBrowserCmd('snapshot', ['-i', '-s', selector]);
        if (result && result.includes('@e')) {
          myClassElements.push({ selector, content: result });
        }
      } catch (e) {
        // 选择器不存在，继续尝试下一个
      }
    }

    console.log(`找到 ${myClassElements.length} 个可能的导航元素`);

    // 6. 检查是否有登录提示或课程列表
    if (snapshot.includes('登录') || snapshot.includes('请登录')) {
      console.log('⚠️ 检测到需要登录');
      results.errors.push('需要登录 - 请先手动登录并保存会话');
      console.log('请手动登录后重新运行自动化');
    } else if (snapshot.includes('我的课堂') || snapshot.includes('课程列表')) {
      console.log('✅ 已登录，开始检查课程进度...');
      
      // 查找"未通过"或"未完成"的课程
      const incompletePatterns = ['未通过', '未完成', '未学完', '进行中'];
      
      for (const pattern of incompletePatterns) {
        if (snapshot.includes(pattern)) {
          console.log(`📚 找到 "${pattern}" 状态的课程`);
          
          // 尝试点击该课程
          try {
            await runAgentBrowserCmd('click', [`text=${pattern}`]);
            await sleep(2000);
            
            // 获取课程详情页
            const lessonSnapshot = await runAgentBrowserCmd('snapshot', ['-i']);
            
            // 查找"听课"按钮
            if (lessonSnapshot.includes('听课')) {
              await runAgentBrowserCmd('click', ['text=听课']);
              await sleep(3000);
              
              // 开始播放
              console.log('▶️ 开始播放课程...');
              await takeScreenshot(`bicpaedu_playing_${timestamp}.png`);
              
              // 等待视频播放一段时间
              await sleep(CONFIG.videoWaitTime);
              
              results.coursesPlayed++;
            }
          } catch (e) {
            console.log(`点击课程失败: ${e.message}`);
            results.errors.push(`点击课程失败: ${e.message}`);
          }
        }
      }
    }

    // 7. 最终截图记录学时统计
    await takeScreenshot(`bicpaedu_end_${timestamp}.png`);
    results.screenshots.push(`bicpaedu_end_${timestamp}.png`);

    // 8. 关闭浏览器
    console.log('\n🔚 关闭浏览器...');
    await runAgentBrowserCmd('close');

    // 9. 输出结果汇总
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    console.log('\n========== 任务执行完成 ==========');
    console.log(`⏱️  总耗时: ${duration}秒`);
    console.log(`📚 完成课程数: ${results.coursesPlayed}`);
    console.log(`📸 截图数量: ${results.screenshots.length}`);
    if (results.errors.length > 0) {
      console.log(`❌ 错误: ${results.errors.length}个`);
      results.errors.forEach((err, i) => console.log(`   ${i + 1}. ${err}`));
    }
    console.log('===================================\n');

    // 返回结果供自动化系统记录
    return {
      success: true,
      coursesPlayed: results.coursesPlayed,
      screenshots: results.screenshots,
      errors: results.errors,
      duration,
    };

  } catch (error) {
    console.error('❌ 执行出错:', error.message);
    results.errors.push(error.message);
    
    try {
      await takeScreenshot(`bicpaedu_error_${new Date().toISOString().slice(0, 10)}.png`);
    } catch (e) {}
    
    await runAgentBrowserCmd('close');
    
    return {
      success: false,
      ...results,
      error: error.message,
    };
  }
}

// 执行主流程
main()
  .then(result => {
    console.log('最终结果:', JSON.stringify(result, null, 2));
    process.exit(result.success ? 0 : 1);
  })
  .catch(err => {
    console.error('未捕获的错误:', err);
    process.exit(1);
  });
