$path = 'server.ts'
$content = Get-Content $path -Raw -Encoding UTF8

$newErrorBlock = @'
  } catch (err: any) {
    console.error('Lỗi API:', err);
    let msg = err?.message || 'Lỗi kết nối máy chủ hoặc API Key không hợp lệ.';
    
    // Check for Quota/Rate Limit errors
    if (msg.includes('429') || msg.toLowerCase().includes('quota') || msg.toLowerCase().includes('exhausted')) {
      msg = 'Hệ thống đang bị quá tải hoặc Key của bạn đã hết lượt tạo miễn phí (Lỗi 429 - Quota Exceeded). Vui lòng thử lại sau khoảng 1-2 phút (để hồi phục token) hoặc sử dụng API Key cá nhân của bạn để tiếp tục.';
    }

    return res.status(500).json({ error: msg });
  }
'@

# Replace in /api/generate
$content = $content -replace '(?s)\} catch \(err: any\) \{\s*console\.error\(''Lỗi khi sinh nội dung:'', err\);\s*const msg = err\?\.message \|\| ''Lỗi kết nối máy chủ hoặc API Key không hợp lệ\.'';\s*return res\.status\(500\)\.json\(\{ error: msg \}\);\s*\}', $newErrorBlock

# Replace in /api/lessons
$content = $content -replace '(?s)\} catch \(err: any\) \{\s*console\.error\(''Lỗi khi tải danh sách bài học:'', err\);\s*return res\.status\(500\)\.json\(\{ error: err\?\.message \|\| ''Không thể tải danh sách bài học\.'' \}\);\s*\}', $newErrorBlock

Set-Content -Path $path -Value $content -Encoding UTF8
