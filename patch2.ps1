$path = 'server.ts'
$content = Get-Content $path -Raw -Encoding UTF8

$newErrorBlock = @'
  } catch (err: any) {
    console.error('Lỗi API:', err);
    let originalMsg = err?.message || 'Lỗi kết nối máy chủ hoặc API Key không hợp lệ.';
    let msg = originalMsg;
    
    // Check for Quota/Rate Limit errors
    if (msg.includes('429') || msg.toLowerCase().includes('quota') || msg.toLowerCase().includes('exhausted')) {
      msg = Hệ thống đang bị quá tải hoặc Key của bạn đã hết lượt tạo miễn phí (Lỗi 429 - Quota Exceeded). Lời nhắn từ máy chủ: "". Vui lòng sử dụng API Key cá nhân của bạn để tiếp tục.;
    }

    return res.status(500).json({ error: msg });
  }
'@

# Replace in /api/generate
$content = $content -replace '(?s)\} catch \(err: any\) \{\s*console\.error\(''Lỗi API:'', err\);\s*let msg = err\?\.message \|\| ''Lỗi kết nối máy chủ hoặc API Key không hợp lệ\.'';\s*// Check for Quota/Rate Limit errors\s*if \(msg\.includes\(''429''\) \|\| msg\.toLowerCase\(\)\.includes\(''quota''\) \|\| msg\.toLowerCase\(\)\.includes\(''exhausted''\)\) \{\s*msg = ''Hệ thống đang bị quá tải hoặc Key của bạn đã hết lượt tạo miễn phí \(Lỗi 429 - Quota Exceeded\)\. Vui lòng thử lại sau khoảng 1-2 phút \(để hồi phục token\) hoặc sử dụng API Key cá nhân của bạn để tiếp tục\.'';\s*\}\s*return res\.status\(500\)\.json\(\{ error: msg \}\);\s*\}', $newErrorBlock

Set-Content -Path $path -Value $content -Encoding UTF8
