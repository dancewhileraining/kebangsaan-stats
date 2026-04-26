<?php
/**
 * Plugin Name: Kebangsaan Statistics
 * Description: Upload Excel per bulan, tampilkan statistik kedatangan & keberangkatan per kebangsaan.
 * Version: 2.4.3
 * Author: Imigrasi Ngurah Rai
 */

if ( ! defined( 'ABSPATH' ) ) exit;

define( 'KBGS_VERSION', '2.4.3' );
define( 'KBGS_PATH',    plugin_dir_path( __FILE__ ) );
define( 'KBGS_URL',     plugin_dir_url( __FILE__ ) );
define( 'KBGS_TABLE',   'kbgs_data' );

/* ── INSTALL ── */
register_activation_hook( __FILE__, 'kbgs_install' );
function kbgs_install() {
    global $wpdb;
    $t  = $wpdb->prefix . KBGS_TABLE;
    $cs = $wpdb->get_charset_collate();
    $sql = "CREATE TABLE IF NOT EXISTS $t (
        id         BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
        year       SMALLINT    NOT NULL,
        month      TINYINT     NOT NULL,
        kebangsaan VARCHAR(160) NOT NULL,
        datang     INT         NOT NULL DEFAULT 0,
        berangkat  INT         NOT NULL DEFAULT 0,
        UNIQUE KEY ym_kb (year, month, kebangsaan)
    ) $cs;";
    require_once ABSPATH . 'wp-admin/includes/upgrade.php';
    dbDelta( $sql );
}

/* ── ADMIN MENU ── */
add_action( 'admin_menu', 'kbgs_admin_menu' );
function kbgs_admin_menu() {
    add_menu_page(
        'Statistik Kebangsaan', 'Statistik Kebangsaan', 'edit_others_posts',
        'kbgs-upload', 'kbgs_upload_page', 'dashicons-chart-bar', 30
    );
    add_submenu_page( 'kbgs-upload', 'Upload Data', 'Upload Data', 'edit_others_posts', 'kbgs-upload', 'kbgs_upload_page' );
    add_submenu_page( 'kbgs-upload', 'Kelola Data', 'Kelola Data', 'edit_others_posts', 'kbgs-manage', 'kbgs_manage_page' );
}

/* ── UPLOAD PAGE ── */
function kbgs_upload_page() {
    $msg = ''; $type = 'info';
    if ( isset($_POST['kbgs_nonce']) && wp_verify_nonce($_POST['kbgs_nonce'], 'kbgs_upload') ) {
        if ( ! empty($_FILES['kbgs_file']['tmp_name']) ) {
            $r = kbgs_process_upload( $_FILES['kbgs_file'] );
            $msg = $r['message']; $type = $r['success'] ? 'success' : 'error';
        } else {
            $msg = 'Pilih file Excel (.xlsx) terlebih dahulu.'; $type = 'error';
        }
    }
    $lb = kbgs_labels();
    ?>
    <div class="wrap">
        <h1>Statistik Kebangsaan — Upload Data</h1>

        <?php if ($msg): ?>
        <div class="notice notice-<?php echo esc_attr($type); ?> is-dismissible"><p><?php echo esc_html($msg); ?></p></div>
        <?php endif; ?>

        <div style="background:#fff;border:1px solid #ccd0d4;border-radius:6px;padding:28px;max-width:620px;margin-top:18px">
            <form method="post" enctype="multipart/form-data" id="kbgs-upload-form">
                <?php wp_nonce_field('kbgs_upload','kbgs_nonce'); ?>
                <table class="form-table">

                    <!-- Upload type -->
                    <tr>
                        <th>Tipe Upload</th>
                        <td>
                            <label>
                                <input type="radio" name="kbgs_upload_type" value="year" checked onchange="kbgsToggleUploadType(this.value)">
                                <strong>Per Tahun</strong> — diambil dari file XLSX Google Sheet Datang dan Berangkat per Kebangsaan yang di-share oleh Bidang TPI
                            </label><br>
                            <label style="margin-top:8px;display:block">
                                <input type="radio" name="kbgs_upload_type" value="month" onchange="kbgsToggleUploadType(this.value)">
                                <strong>Per Bulan</strong> — diambil dari file XLSX Statistik Datang dan Berangkat yang dikirim melalui email Tikkim oleh Bidang TPI
                            </label>
                        </td>
                    </tr>

                    <!-- File -->
                    <tr>
                        <th><label for="kbgs_file">File Excel</label></th>
                        <td>
                            <input type="file" name="kbgs_file" id="kbgs_file" accept=".xlsx" required style="width:100%">
                            <p class="description">Format .xlsx — maks. 10 MB</p>
                        </td>
                    </tr>

                    <!-- Year -->
                    <tr>
                        <th><label for="kbgs_year">Tahun</label></th>
                        <td><input type="number" name="kbgs_year" id="kbgs_year" value="" placeholder="e.g. 2025" min="2020" max="2099" class="small-text" required></td>
                    </tr>

                    <!-- Month (only shown for single-month upload) -->
                    <tr id="kbgs-row-month" style="display:none">
                        <th><label for="kbgs_month">Bulan</label></th>
                        <td>
                            <select name="kbgs_month" id="kbgs_month" required>
                                <option value="">— Pilih Bulan —</option>
                                <?php foreach ($lb as $num => $name): ?>
                                <option value="<?php echo $num; ?>"><?php echo esc_html($name); ?></option>
                                <?php endforeach; ?>
                            </select>
                        </td>
                    </tr>

                    <!-- Mode -->
                    <tr>
                        <th>Mode Simpan</th>
                        <td>
                            <label><input type="radio" name="kbgs_mode" value="replace" checked> <strong>Ganti</strong> — hapus data yang ada lalu isi ulang</label><br>
                            <label style="margin-top:6px;display:block"><input type="radio" name="kbgs_mode" value="merge"> <strong>Gabung</strong> — tambah/update tanpa hapus bulan lain</label>
                        </td>
                    </tr>
                </table>
                <p style="margin-top:20px">
                    <button type="submit" class="button button-primary button-large">Upload &amp; Proses</button>
                </p>
            </form>
        </div>

        <!-- Format hints -->
        <div style="background:#f0f6fc;border:1px solid #c3d9f0;border-radius:6px;padding:20px;max-width:620px;margin-top:16px">
            <h3 style="margin-top:0">Format File</h3>
            <p><strong>Per Tahun:</strong> sheet dinamai <code>Jan, Feb, Mar, Apr, Mei, Jun, Jul, Agu, Sep, Okt, Nov, Des</code></p>
            <p style="margin-bottom:0"><strong>Per Bulan:</strong> sheet bernama <code>REKAP</code> — Kebangsaan C13:C282, Kedatangan I13:I282, Keberangkatan J13:J282</p>
        </div>

        <div style="background:#f0f6fc;border:1px solid #c3d9f0;border-radius:6px;padding:20px;max-width:620px;margin-top:16px">
            <h3 style="margin-top:0">Shortcode</h3>
            <code style="display:block;padding:10px;background:#fff;border:1px solid #ddd;border-radius:4px;font-size:14px">[kebangsaan_stats]</code>
            <p style="margin-bottom:0;margin-top:8px">Opsional: <code>[kebangsaan_stats year="2026" default_month="1"]</code></p>
        </div>
    </div>

    <script>
    function kbgsToggleUploadType(val) {
        document.getElementById('kbgs-row-month').style.display = (val === 'month') ? '' : 'none';
    }
    </script>
    <?php
}

/* ── MANAGE PAGE ── */
function kbgs_manage_page() {
    global $wpdb;
    $t = $wpdb->prefix . KBGS_TABLE;
    if ( isset($_POST['kbgs_del_nonce']) && wp_verify_nonce($_POST['kbgs_del_nonce'],'kbgs_delete') ) {
        $year = intval($_POST['del_year']); $month = intval($_POST['del_month']);
        if ($month === 0) {
            $wpdb->delete($t,['year'=>$year],['%d']);
            echo '<div class="notice notice-success is-dismissible"><p>Data tahun '.$year.' dihapus.</p></div>';
        } else {
            $wpdb->delete($t,['year'=>$year,'month'=>$month],['%d','%d']);
            $lb = kbgs_labels();
            echo '<div class="notice notice-success is-dismissible"><p>Data '.esc_html($lb[$month]).' '.$year.' dihapus.</p></div>';
        }
    }
    $rows = $wpdb->get_results("SELECT year,month,COUNT(*) cnt,SUM(datang) td,SUM(berangkat) tb FROM {$t} GROUP BY year,month ORDER BY year DESC,month ASC");
    $lb   = kbgs_labels();
    ?>
    <div class="wrap"><h1>Kelola Data Tersimpan</h1>
    <?php if (empty($rows)): ?>
        <p>Belum ada data. <a href="<?php echo admin_url('admin.php?page=kbgs-upload'); ?>">Upload sekarang</a>.</p>
    <?php else: ?>
    <table class="wp-list-table widefat fixed striped" style="max-width:820px">
        <thead><tr><th>Tahun</th><th>Bulan</th><th>Negara</th><th>Total Datang</th><th>Total Berangkat</th><th>Aksi</th></tr></thead>
        <tbody>
        <?php foreach ($rows as $r): ?>
        <tr>
            <td><?php echo esc_html($r->year); ?></td>
            <td><?php echo esc_html($lb[$r->month] ?? $r->month); ?></td>
            <td><?php echo number_format($r->cnt); ?></td>
            <td><?php echo number_format($r->td); ?></td>
            <td><?php echo number_format($r->tb); ?></td>
            <td>
                <form method="post" style="display:inline" onsubmit="return confirm('Hapus data ini?')">
                    <?php wp_nonce_field('kbgs_delete','kbgs_del_nonce'); ?>
                    <input type="hidden" name="del_year"  value="<?php echo esc_attr($r->year); ?>">
                    <input type="hidden" name="del_month" value="<?php echo esc_attr($r->month); ?>">
                    <button type="submit" class="button button-small button-link-delete">Hapus</button>
                </form>
            </td>
        </tr>
        <?php endforeach; ?>
        </tbody>
    </table>
    <?php endif; ?></div>
    <?php
}

/* ── PROCESS UPLOAD ── */
function kbgs_process_upload( $file ) {
    $year        = intval($_POST['kbgs_year'] ?? date('Y'));
    $mode        = sanitize_text_field($_POST['kbgs_mode'] ?? 'replace');
    $upload_type = sanitize_text_field($_POST['kbgs_upload_type'] ?? 'year');

    if ( strtolower(pathinfo($file['name'], PATHINFO_EXTENSION)) !== 'xlsx' )
        return ['success'=>false,'message'=>'Hanya file .xlsx yang didukung.'];

    global $wpdb;
    $t  = $wpdb->prefix . KBGS_TABLE;
    $lb = kbgs_labels();

    /* ── Single-month upload ── */
    if ( $upload_type === 'month' ) {
        $month = intval($_POST['kbgs_month'] ?? date('n'));
        if ( $month < 1 || $month > 12 )
            return ['success'=>false,'message'=>'Pilih bulan yang valid.'];

        $rows = kbgs_parse_xlsx_single($file['tmp_name']);
        if ( empty($rows) )
            return ['success'=>false,'message'=>'Tidak ada data terbaca dari file. Pastikan header KEBANGSAAN / KEDATANGAN / KEBERANGKATAN ada di sheet.'];

        $wpdb->delete($t, ['year'=>$year,'month'=>$month], ['%d','%d']);
        foreach ($rows as $row) {
            $wpdb->replace($t,[
                'year'=>$year,'month'=>$month,
                'kebangsaan'=>$row['kebangsaan'],
                'datang'=>$row['datang'],
                'berangkat'=>$row['berangkat'],
            ],['%d','%d','%s','%d','%d']);
        }
        $mn = $lb[$month] ?? $month;
        return ['success'=>true,'message'=>"Berhasil! ".count($rows)." baris disimpan untuk {$mn} {$year}."];
    }

    /* ── Full-year upload (existing behaviour) ── */
    $data = kbgs_parse_xlsx($file['tmp_name']);
    if ( empty($data) )
        return ['success'=>false,'message'=>'Tidak ada data terbaca. Pastikan nama sheet sesuai (Jan, Feb, Mar, ...).'];

    if ($mode === 'replace') $wpdb->delete($t,['year'=>$year],['%d']);
    $inserted = 0;
    foreach ($data as $m => $rows) {
        if ($mode === 'merge') $wpdb->delete($t,['year'=>$year,'month'=>$m],['%d','%d']);
        foreach ($rows as $row) {
            $wpdb->replace($t,[
                'year'=>$year,'month'=>$m,
                'kebangsaan'=>$row['kebangsaan'],
                'datang'=>$row['datang'],
                'berangkat'=>$row['berangkat'],
            ],['%d','%d','%s','%d','%d']);
            $inserted++;
        }
    }
    return ['success'=>true,'message'=>"Berhasil! {$inserted} baris dari ".count($data)." bulan disimpan untuk tahun {$year}."];
}

/* ── XLSX PARSER (pure PHP / ZipArchive) ──
 *
 * KEY FIX: Self-closing empty cells like <c r="D14" s="32"/> were causing the
 * greedy (.*?) regex to consume everything up to the NEXT </c>, swallowing
 * columns E-I into column D. Fix: strip self-closing cells first, then parse.
 *
 * Also uses relationship file to correctly map sheet names to XML files,
 * regardless of file numbering order.
 * ─────────────────────────────────────────── */
function kbgs_parse_xlsx( $path ) {
    $month_map = [
        'jan'=>1,'feb'=>2,'mar'=>3,'apr'=>4,'mei'=>5,'may'=>5,
        'jun'=>6,'jul'=>7,'agu'=>8,'aug'=>8,'sep'=>9,
        'okt'=>10,'oct'=>10,'nov'=>11,'des'=>12,'dec'=>12
    ];

    $zip = new ZipArchive();
    if ($zip->open($path) !== true) return [];

    // 1. Shared strings — parse per <si> to handle rich-text entries correctly
    $strings = [];
    $sst = $zip->getFromName('xl/sharedStrings.xml');
    if ($sst) {
        preg_match_all('/<si>(.*?)<\/si>/s', $sst, $si_m);
        foreach ($si_m[1] as $si_content) {
            preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $si_content, $t_m);
            $strings[] = html_entity_decode(implode('', $t_m[1]));
        }
    }

    // 2. Relationship file: rId -> worksheet file path
    $rid_to_file = [];
    $rels_xml = $zip->getFromName('xl/_rels/workbook.xml.rels');
    if ($rels_xml) {
        preg_match_all('/Id="(rId\d+)"[^>]*Target="(worksheets\/sheet\d+\.xml)"/', $rels_xml, $rm);
        foreach ($rm[1] as $i => $rid) {
            $rid_to_file[$rid] = $rm[2][$i];
        }
    }

    // 3. Workbook: sheet name -> rId -> actual file
    $sheet_files = []; // sheet_name (lowercase) => 'worksheets/sheetN.xml'
    $wb = $zip->getFromName('xl/workbook.xml');
    if ($wb) {
        preg_match_all('/name="([^"]+)"[^>]+r:id="(rId\d+)"/', $wb, $sm);
        foreach ($sm[1] as $i => $name) {
            $rid  = $sm[2][$i];
            $file = $rid_to_file[$rid] ?? null;
            if ($file) $sheet_files[strtolower(trim($name))] = $file;
        }
    }

    // 4. Parse each monthly sheet
    $result = [];
    foreach ($month_map as $sname => $mnum) {
        if (isset($result[$mnum])) continue; // already parsed this month
        $file = $sheet_files[$sname] ?? null;
        if (!$file) continue;
        $xml = $zip->getFromName('xl/' . $file);
        if (!$xml) continue;
        $rows = kbgs_parse_sheet($xml, $strings);
        if (!empty($rows)) $result[$mnum] = $rows;
    }

    $zip->close();
    return $result;
}

/* ── SINGLE-SHEET PARSER ──
 * Finds the right summary sheet in a single-month xlsx, then parses it.
 * Strategy:
 *   1. Prefer a sheet whose name is a month abbreviation (Jan/Des/DES/…)
 *   2. Fall back to trying every non-system sheet until we get rows
 *   3. Last resort: first sheet
 */
function kbgs_parse_xlsx_single( $path ) {
    $zip = new ZipArchive();
    if ($zip->open($path) !== true) return [];

    // Shared strings — parse per <si> to handle rich-text entries correctly
    $strings = [];
    $sst = $zip->getFromName('xl/sharedStrings.xml');
    if ($sst) {
        preg_match_all('/<si>(.*?)<\/si>/s', $sst, $si_m);
        foreach ($si_m[1] as $si_content) {
            preg_match_all('/<t[^>]*>(.*?)<\/t>/s', $si_content, $t_m);
            $strings[] = html_entity_decode(implode('', $t_m[1]));
        }
    }

    // Build rId → file map
    $rid_map = [];
    $rels_xml = $zip->getFromName('xl/_rels/workbook.xml.rels');
    if ($rels_xml) {
        preg_match_all('/Id="(rId\d+)"[^>]*Target="(worksheets\/sheet\d+\.xml)"/', $rels_xml, $rm);
        foreach ($rm[1] as $i => $rid) $rid_map[$rid] = $rm[2][$i];
    }

    // Build ordered list: sheet name → file, in workbook order
    $sheets_ordered = [];
    $wb = $zip->getFromName('xl/workbook.xml');
    if ($wb) {
        preg_match_all('/name="([^"]+)"[^>]+r:id="(rId\d+)"/', $wb, $sm);
        foreach ($sm[1] as $i => $name) {
            $rid  = $sm[2][$i];
            $file = $rid_map[$rid] ?? null;
            if ($file) $sheets_ordered[$name] = $file;
        }
    }

    // 1. Prefer "REKAP" — use dedicated hardcoded-range parser (C13:C282, I13:I282, J13:J282)
    $target_file     = null;
    $use_rekap_parser = false;
    foreach ($sheets_ordered as $name => $file) {
        if (strtolower(trim($name)) === 'rekap') {
            $target_file      = $file;
            $use_rekap_parser = true;
            break;
        }
    }

    // 2. Try every non-system sheet in order until we get valid rows
    $system_prefixes = ['_xlnm', 'Z_A', 'microsoft', 'Filter'];
    if (!$target_file) {
        foreach ($sheets_ordered as $name => $file) {
            $skip = false;
            foreach ($system_prefixes as $pfx) {
                if (stripos($name, $pfx) === 0) { $skip = true; break; }
            }
            if ($skip) continue;
            $xml = $zip->getFromName('xl/' . $file);
            if (!$xml) continue;
            $rows = kbgs_parse_sheet($xml, $strings);
            if (!empty($rows)) { $zip->close(); return $rows; }
        }
    }

    // 3. Last resort: first sheet
    if (!$target_file) $target_file = reset($sheets_ordered) ?: 'worksheets/sheet1.xml';

    $xml = $zip->getFromName('xl/' . $target_file);
    $zip->close();
    if (!$xml) return [];
    return $use_rekap_parser
        ? kbgs_parse_rekap_sheet($xml, $strings)
        : kbgs_parse_sheet($xml, $strings);
}

/* ── REKAP SHEET PARSER ──
 * Hardcoded cell ranges per user spec:
 *   C13:C282  = Kebangsaan
 *   I13:I282  = Kedatangan
 *   J13:J282  = Keberangkatan
 *   I283      = Total Kedatangan
 *   J283      = Total Keberangkatan  (read but not stored — DB totals are computed)
 */
function kbgs_parse_rekap_sheet( $xml, $strings ) {
    $grid = [];

    preg_match_all('/<row[^>]+r="(\d+)"[^>]*>(.*?)<\/row>/s', $xml, $rm);
    foreach ($rm[1] as $idx => $row_num) {
        $rn      = (int) $row_num;
        if ($rn < 13 || $rn > 283) continue;   // only rows we care about
        $row_xml = $rm[2][$idx];
        $row_xml = preg_replace('/<c\s[^>]*\/>/s', '', $row_xml);
        preg_match_all('/<c\s+r="([A-Z]+)(\d+)"([^>]*)>(.*?)<\/c>/s', $row_xml, $cm, PREG_SET_ORDER);
        foreach ($cm as $cell) {
            $col  = $cell[1];
            if (!in_array($col, ['C','I','J'])) continue;
            $attrs = $cell[3];
            $inner = $cell[4];
            $is_s  = (bool) preg_match('/\bt="s"/', $attrs);
            preg_match('/<v>(.*?)<\/v>/s', $inner, $vm);
            $val = isset($vm[1]) ? trim($vm[1]) : '';
            if ($is_s && isset($strings[(int)$val])) {
                $val = $strings[(int)$val];
            }
            $grid[$rn][$col] = trim(strip_tags($val));
        }
    }

    $out = [];
    for ($r = 13; $r <= 282; $r++) {
        $kb = trim($grid[$r]['C'] ?? '');
        if (!$kb || is_numeric($kb)) continue;
        $up = strtoupper($kb);
        if (in_array($up, ['KEBANGSAAN','NO','JUMLAH','TOTAL'])) continue;
        $dt = $grid[$r]['I'] ?? '0';
        $br = $grid[$r]['J'] ?? '0';
        $out[] = [
            'kebangsaan' => strtoupper($kb),
            'datang'     => intval(floatval($dt)),
            'berangkat'  => intval(floatval($br)),
        ];
    }
    return $out;
}

function kbgs_parse_sheet( $xml, $strings ) {
    $grid = [];

    // Parse row by row
    preg_match_all('/<row[^>]+r="(\d+)"[^>]*>(.*?)<\/row>/s', $xml, $rm);

    foreach ($rm[1] as $idx => $row_num) {
        $row_xml = $rm[2][$idx];

        // ── CRITICAL FIX ──
        // Remove self-closing empty cells <c r="X99" .../> BEFORE running cell regex.
        // Without this, the (.*?) in the cell regex greedily consumes the text of
        // following cells until it finds a </c>, corrupting all column assignments.
        $row_xml = preg_replace('/<c\s[^>]*\/>/s', '', $row_xml);

        // Now safely parse non-empty cells
        preg_match_all('/<c\s+r="([A-Z]+)(\d+)"([^>]*)>(.*?)<\/c>/s', $row_xml, $cm, PREG_SET_ORDER);

        foreach ($cm as $cell) {
            $col   = $cell[1];
            $attrs = $cell[3];
            $inner = $cell[4];
            $is_s  = (bool) preg_match('/\bt="s"/', $attrs);
            preg_match('/<v>(.*?)<\/v>/s', $inner, $vm);
            $val = isset($vm[1]) ? trim($vm[1]) : '';
            if ($is_s && isset($strings[(int)$val])) {
                $val = $strings[(int)$val];
            }
            $grid[(int)$row_num][$col] = trim(strip_tags($val));
        }
    }

    // Fixed column positions for this Excel format:
    // B = row number, C = country, I = datang, J = berangkat
    // Data rows: 14 to 283 (skip header rows and column-label row 13)
    $col_kb = 'C';
    $col_dt = 'I';
    $col_br = 'J';

    // Auto-detect: scan for the header row containing 'KEBANGSAAN'
    // then look for KEDATANGAN/KEBERANGKATAN in surrounding rows
    $header_row = 0;
    foreach ($grid as $rn => $cols) {
        foreach ($cols as $col => $val) {
            if (strtoupper(trim($val)) === 'KEBANGSAAN') {
                $col_kb = $col;
                $header_row = $rn;
                break 2;
            }
        }
    }

    // Scan rows near header for KEDATANGAN / KEBERANGKATAN
    if ($header_row) {
        for ($r = $header_row; $r <= $header_row + 5; $r++) {
            if (!isset($grid[$r])) continue;
            foreach ($grid[$r] as $col => $val) {
                $up = strtoupper(trim($val));
                if ($up === 'KEDATANGAN'  && !$col_dt) $col_dt = $col;
                if ($up === 'KEBERANGKATAN' && !$col_br) $col_br = $col;
            }
        }
        // If both landed on the same column (merged-cell header edge case),
        // shift keberangkatan to the next column.
        if ($col_br && $col_dt && $col_br === $col_dt) {
            $col_br = chr(ord($col_dt) + 1);
        }
    }

    // Fix col_kb: KEBANGSAAN label may be one column LEFT of actual country data
    // (e.g. header at B but country names are at C). Verify against first data row.
    if ($header_row && $col_kb) {
        foreach ($grid as $rn => $cols) {
            if ($rn <= $header_row + 3) continue;
            $v = trim($cols[$col_kb] ?? '');
            // If col_kb has a number instead of a country name, try the next column
            if ($v !== '' && is_numeric($v)) {
                $next = chr(ord($col_kb) + 1);
                $v2   = trim($cols[$next] ?? '');
                if ($v2 !== '' && !is_numeric($v2)) $col_kb = $next;
            }
            break;
        }
    }

    // Find first real data row: col_kb = non-numeric string, col_dt/col_br = numeric
    $data_start = 0;
    foreach ($grid as $rn => $cols) {
        if ($header_row && $rn <= $header_row + 3) continue; // skip header area
        $kb = trim($cols[$col_kb] ?? '');
        $dt = trim($cols[$col_dt] ?? '');
        $br = trim($cols[$col_br] ?? '');
        if (!$kb || is_numeric($kb)) continue;
        $up = strtoupper($kb);
        if (in_array($up, ['KEBANGSAAN','NO','JUMLAH'])) continue;
        if (is_numeric($dt) || is_numeric($br)) { $data_start = $rn; break; }
    }

    if (!$data_start) return [];

    $out = [];
    foreach ($grid as $rn => $cols) {
        if ($rn < $data_start) continue;
        $kb = trim($cols[$col_kb] ?? '');
        $dt = trim($cols[$col_dt] ?? '0');
        $br = trim($cols[$col_br] ?? '0');
        if (!$kb || is_numeric($kb)) continue;
        $up = strtoupper($kb);
        if (in_array($up, ['JUMLAH','KEBANGSAAN','NO'])) continue;
        $out[] = [
            'kebangsaan' => strtoupper($kb),
            'datang'     => intval(floatval($dt)),
            'berangkat'  => intval(floatval($br)),
        ];
    }
    return $out;
}

/* ── SHORTCODE ── */
add_shortcode( 'kebangsaan_stats', 'kbgs_shortcode' );
function kbgs_shortcode( $atts ) {
    $atts = shortcode_atts([
        'year'          => null,               // lock to a single year (optional)
        'default_year'  => null,               // which year tab opens first
        'default_month' => intval(date('n')),  // which month tab opens first
    ], $atts);

    global $wpdb;
    $t      = $wpdb->prefix . KBGS_TABLE;
    $labels = kbgs_labels();
    $defm   = intval($atts['default_month']);

    /* ── Fetch available years ── */
    if ($atts['year']) {
        $years = [intval($atts['year'])];
    } else {
        $years = array_map('intval',
            $wpdb->get_col("SELECT DISTINCT year FROM {$t} ORDER BY year DESC")
        );
    }
    if (empty($years))
        return '<p style="color:#888;font-style:italic">Data statistik belum tersedia.</p>';

    $defy = $atts['default_year'] ? intval($atts['default_year']) : $years[0];
    if (!in_array($defy, $years)) $defy = $years[0];

    /* ── Fetch all data in one query ── */
    $placeholders = implode(',', array_fill(0, count($years), '%d'));
    $all = $wpdb->get_results(
        $wpdb->prepare(
            "SELECT year,month,kebangsaan,datang,berangkat FROM {$t}
             WHERE year IN ($placeholders) ORDER BY year DESC, month ASC, datang DESC",
            ...$years
        )
    );
    $by = [];  /* $by[year][month][] = row */
    foreach ($all as $row) $by[$row->year][$row->month][] = $row;

    $uid = 'kbgs_' . uniqid();

    wp_enqueue_style(  'kbgs-style',     KBGS_URL . 'assets/style.css', [], KBGS_VERSION );
    wp_enqueue_script( 'kbgs-xlsx',      'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js', [], '0.20.3', true );
    wp_enqueue_script( 'kbgs-jspdf',     'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', [], '2.5.1', true );
    wp_enqueue_script( 'kbgs-autotable', 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js', ['kbgs-jspdf'], '3.8.2', true );
    wp_enqueue_script( 'kbgs-script',    KBGS_URL . 'assets/script.js', ['kbgs-xlsx','kbgs-autotable'], KBGS_VERSION, true );

    ob_start();
    ?>
    <div class="kbgs-root" id="<?php echo esc_attr($uid); ?>">

        <!-- Top bar: year tabs left, search right -->
        <div class="kbgs-topbar">
            <div class="kbgs-year-tabs" role="tablist" aria-label="Pilih tahun">
            <?php foreach ($years as $y): ?>
            <button
                class="kbgs-year-tab <?php echo $y==$defy?'active':''; ?>"
                data-year="<?php echo esc_attr($y); ?>"
                role="tab"
                aria-selected="<?php echo $y==$defy?'true':'false'; ?>"
                onclick="kbgsYearTab(this,'<?php echo esc_attr($uid); ?>')">
                <?php echo esc_html($y); ?>
            </button>
            <?php endforeach; ?>
            </div><!-- /.kbgs-year-tabs -->
            <input type="text" class="kbgs-search" placeholder="Cari negara..." oninput="kbgsFilter(this)" aria-label="Cari negara">
        </div><!-- /.kbgs-topbar -->

        <!-- Year panels -->
        <?php foreach ($years as $y):
            $avail_m = array_keys($by[$y] ?? []);
            sort($avail_m);
            $active_m = in_array($defm, $avail_m) ? $defm : ($avail_m[0] ?? 1);
        ?>
        <div class="kbgs-year-panel <?php echo $y==$defy?'active':''; ?>" data-year="<?php echo esc_attr($y); ?>">

            <!-- Month tabs -->
            <div class="kbgs-tabs" role="tablist">
                <?php foreach ($avail_m as $m): ?>
                <button
                    class="kbgs-tab <?php echo $m==$active_m?'active':''; ?>"
                    role="tab"
                    data-month="<?php echo esc_attr($m); ?>"
                    aria-selected="<?php echo $m==$active_m?'true':'false'; ?>"
                    onclick="kbgsTab(this,'<?php echo esc_attr($uid); ?>')">
                    <?php echo esc_html($labels[$m] ?? $m); ?>
                </button>
                <?php endforeach; ?>
            </div>

            <!-- Month panels -->
            <?php foreach ($avail_m as $m):
                $rows  = $by[$y][$m] ?? [];
                $tot_d = array_sum(array_column($rows, 'datang'));
                $tot_b = array_sum(array_column($rows, 'berangkat'));
            ?>
            <div class="kbgs-panel <?php echo $m==$active_m?'active':''; ?>" data-month="<?php echo esc_attr($m); ?>">

                <!-- Dashboard -->
                <div class="kbgs-dashboard">
                    <div class="kbgs-stat kbgs-stat-datang">
                        <span class="kbgs-stat-label">Total Kedatangan</span>
                        <span class="kbgs-stat-value"><?php echo number_format($tot_d); ?></span>
                        <span class="kbgs-stat-sub"><?php echo esc_html($labels[$m].' '.$y); ?></span>
                    </div>
                    <div class="kbgs-stat kbgs-stat-berangkat">
                        <span class="kbgs-stat-label">Total Keberangkatan</span>
                        <span class="kbgs-stat-value"><?php echo number_format($tot_b); ?></span>
                        <span class="kbgs-stat-sub"><?php echo esc_html($labels[$m].' '.$y); ?></span>
                    </div>
                </div>

                <!-- Table -->
                <div class="kbgs-table-wrap">
                    <table class="kbgs-table">
                        <colgroup>
                            <col style="width:64px">
                            <col style="width:38%">
                            <col>
                            <col>
                        </colgroup>
                        <thead>
                            <tr>
                                <th class="kbgs-th-no">NO</th>
                                <th class="kbgs-th-kb">KEBANGSAAN</th>
                                <th class="kbgs-th-datang" onclick="kbgsSort(this,'kbgs-num-datang')">KEDATANGAN <span class="kbgs-si">↕</span></th>
                                <th class="kbgs-th-berangkat" onclick="kbgsSort(this,'kbgs-num-berangkat')">KEBERANGKATAN <span class="kbgs-si">↕</span></th>
                            </tr>
                        </thead>
                        <tbody>
                        <?php foreach ($rows as $i => $row): ?>
                            <tr>
                                <td class="kbgs-td-no"><?php echo $i+1; ?></td>
                                <td class="kbgs-td-kb"><?php echo esc_html(ucwords(strtolower($row->kebangsaan))); ?></td>
                                <td class="kbgs-num-datang"><?php echo number_format($row->datang); ?></td>
                                <td class="kbgs-num-berangkat"><?php echo number_format($row->berangkat); ?></td>
                            </tr>
                        <?php endforeach; ?>
                        </tbody>
                        <tfoot>
                            <tr>
                                <td colspan="2" class="kbgs-foot-label">JUMLAH</td>
                                <td class="kbgs-num-datang kbgs-foot-num"><?php echo number_format($tot_d); ?></td>
                                <td class="kbgs-num-berangkat kbgs-foot-num"><?php echo number_format($tot_b); ?></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>

                <!-- Export bar (below table) -->
                <div class="kbgs-export-bar">
                    <button class="kbgs-btn-export kbgs-btn-excel"
                        onclick="kbgsExportExcel(this)"
                        data-month-label="<?php echo esc_attr($labels[$m]); ?>"
                        data-year="<?php echo esc_attr($y); ?>">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
                        Download Excel
                    </button>
                    <button class="kbgs-btn-export kbgs-btn-pdf"
                        onclick="kbgsExportPdf(this)"
                        data-month-label="<?php echo esc_attr($labels[$m]); ?>"
                        data-year="<?php echo esc_attr($y); ?>">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
                        Download PDF
                    </button>
                </div>

            </div>
            <?php endforeach; ?>

        </div><!-- .kbgs-year-panel -->
        <?php endforeach; ?>

    </div><!-- .kbgs-root -->
    <?php
    return ob_get_clean();
}

/* ── HELPERS ── */
function kbgs_labels() {
    return [1=>'Januari',2=>'Februari',3=>'Maret',4=>'April',5=>'Mei',6=>'Juni',
            7=>'Juli',8=>'Agustus',9=>'September',10=>'Oktober',11=>'November',12=>'Desember'];
}
