/**
 * CPA继续教育自动播放脚本 v2.0
 * 自动化播放北京市注协CPA继续教育课程
 * 
 * 使用方法：
 * 1. 首次运行：脚本会自动登录（如果需要）
 * 2. 后续运行：使用保存的会话自动继续播放
 */

const { exec } = require('child_process');
const path = require('path');

// 配置
const CONFIG = {
  // 登录凭证
  username: '1120260109',
  password: 'suhao',
  
  // 会话名称
  sessionName: 'bicpaedu',
  
  // 截图目录
  screenshotDir: 'c:\\Users\\12844\\WorkBuddy\\Claw\\screenshots',
  
  // 播放设置
  maxLessonsPerRun: 3,        // 每次运行最多播放3节课
  waitTimeBetweenLessons: 5000, // 课程之间等待时间(ms)
  videoCheckInterval: 30000,   // 检查视频播放状态间隔(ms)
  
  // URLs
  baseUrl: 'https://www.bicpaedu.com',
  loginUrl: 'https://www.bicpaedu.com/bicpa/mem/mem_toLogin.jspr',
  courseListUrl: 'https://www.bicpaedu.com/bicpa/edu/learnUserCourse_indexJXJY.jspr',
};

// 辅助函数：执行agent-browser命令
function runAgentBrowserCmd(cmd, args = []) {
  return new Promise((resolve, reject) => {
    const fullCmd = `agent-browser ${cmd} ${args.join(' ')}`;
    exec(fullCmd, { maxBuffer: 1024 * 1024, cwd: 'c:\\Users\\12844\\WorkBuddy\\Claw' }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(`命令执行失败: ${fullCmd}\n错误: ${error.message}`));
        return;
      }
      resolve(stdout);
    });
  });
}

// 辅助函数：等待
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 截图并保存
async function takeScreenshot(filename) {
  try {
    const screenshotPath = path.join(CONFIG.screenshotDir, filename);
    await runAgentBrowserCmd('screenshot', [screenshotPath]);
    console.log(`  ✓ 截图: ${filename}`);
    return screenshotPath;
  } catch (error) {
    console.log(`  ⚠️ 截图失败: ${error.message}`);
    return null;
  }
}

// 获取课程列表
async function getCourseList() {
  try {
    await runAgentBrowserCmd('open', [CONFIG.courseListUrl, '--session-name', CONFIG.sessionName]);
    await sleep(2000);
    
    const result = await runAgentBrowserCmd('eval', [
      'JSON.stringify(Array.from(document.querySelectorAll("a")).filter(a => a.href && a.href.includes("learnUserCourse_toLearnChapter")).map(a => ({href: a.href, text: a.textContent.trim() || "听课"})))'
    ]);
    
    return JSON.parse(result || '[]');
  } catch (error) {
    console.log('  ⚠️ 获取课程列表失败');
    return [];
  }
}

// 播放单个课程
async function playCourse(courseUrl, courseName) {
  try {
    console.log(`\n  📚 开始播放: ${courseName || '课程'}`);
    
    await runAgentBrowserCmd('open', [courseUrl, '--session-name', CONFIG.sessionName]);
    await sleep(2000);
    
    // 获取课程小节列表
    const chaptersResult = await runAgentBrowserCmd('eval', [
      'JSON.stringify(Array.from(document.querySelectorAll("a")).filter(a => a.href && a.href.includes("learn_studyOnLine")).map(a => ({href: a.href, text: a.textContent.trim()})))'
    ]);
    
    const chapters = JSON.parse(chaptersResult || '[]');
    console.log(`  📖 找到 ${chapters.length} 个小节`);
    
    // 播放每个小节
    for (let i = 0; i < chapters.length; i++) {
      const chapter = chapters[i];
      console.log(`  ▶️ 播放第 ${i + 1}/${chapters.length} 节: ${chapter.text}`);
      
      await runAgentBrowserCmd('open', [chapter.href, '--session-name', CONFIG.sessionName]);
      await sleep(3000);
      
      // 点击播放按钮
      await runAgentBrowserCmd('eval', [
        'const playBtn = document.querySelector(".play-btn, .vjs-play-control, [class*=play], video"); if(playBtn) { if(playBtn.tagName === "VIDEO") playBtn.play(); else playBtn.click(); }'
      ]);
      
      // 等待视频播放
      await sleep(CONFIG.videoCheckInterval);
      
      // 检查是否播放完成
      const completed = await runAgentBrowserCmd('eval', [
        'const video = document.querySelector("video"); video ? (video.currentTime / video.duration > 0.9) : true'
      ]);
      
      if (completed.includes('true')) {
        console.log(`  ✅ 第 ${i + 1} 节播放完成`);
      }
    }
    
    return true;
  } catch (error) {
    console.log(`  ❌ 播放失败: ${error.message}`);
    return false;
  }
}

// 主流程
async function main() {
  const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  const results = {
    coursesPlayed: 0,
    errors: [],
    screenshots: [],
  };

  console.log('\n========================================');
  console.log('  CPA继续教育自动播放任务');
  console.log('  时间: ' + new Date().toLocaleString('zh-CN'));
  console.log('========================================\n');

  try {
    // 1. 打开浏览器
    console.log('1️⃣ 打开浏览器...');
    await runAgentBrowserCmd('open', [CONFIG.courseListUrl, '--session-name', CONFIG.sessionName]);
    await sleep(3000);
    
    // 2. 检查是否需要登录
    const pageContent = await runAgentBrowserCmd('snapshot', ['-i']);
    
    if (pageContent.includes('请输入会员证书编号')) {
      console.log('📝 需要登录...');
      await runAgentBrowserCmd('fill', ['@e11', CONFIG.username]);
      await runAgentBrowserCmd('fill', ['@e12', CONFIG.password]);
      await runAgentBrowserCmd('click', ['@e5']);
      await sleep(5000);
    }
    
    // 3. 截图记录开始状态
    await takeScreenshot(`bicpaedu_start_${timestamp}.png`);
    
    // 4. 获取课程列表
    console.log('\n2️⃣ 获取课程列表...');
    const courses = await getCourseList();
    console.log(`   找到 ${courses.length} 门课程`);
    
    // 5. 播放课程
    console.log('\n3️⃣ 开始播放课程...');
    for (let i = 0; i < Math.min(courses.length, CONFIG.maxLessonsPerRun); i++) {
      const course = courses[i];
      const success = await playCourse(course.href, course.text);
      if (success) {
        results.coursesPlayed++;
      } else {
        results.errors.push(`播放失败: ${course.text}`);
      }
      
      await sleep(CONFIG.waitTimeBetweenLessons);
    }
    
    // 6. 最终截图
    console.log('\n4️⃣ 截图记录...');
    await takeScreenshot(`bicpaedu_end_${timestamp}.png`);
    
    // 7. 关闭浏览器（保存会话）
    console.log('\n5️⃣ 关闭浏览器...');
    await runAgentBrowserCmd('close');
    
    // 8. 输出结果
    console.log('\n========================================');
    console.log('  任务执行完成');
    console.log(`  播放课程数: ${results.coursesPlayed}`);
    console.log(`  截图数量: ${results.screenshots.length}`);
    if (results.errors.length > 0) {
      console.log(`  错误: ${results.errors.length}个`);
    }
    console.log('========================================\n');

    return { success: true, ...results };

  } catch (error) {
    console.error('❌ 执行出错:', error.message);
    await takeScreenshot(`bicpaedu_error_${timestamp}.png`);
    await runAgentBrowserCmd('close');
    
    return { success: false, error: error.message };
  }
}

// 执行
main().then(result => {
  console.log('结果:', JSON.stringify(result, null, 2));
  process.exit(result.success ? 0 : 1);
}).catch(err => {
  console.error('未捕获错误:', err);
  process.exit(1);
});
