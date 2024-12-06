<?php
require_once plugin_dir_path(__FILE__) . 'excel/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

//echo count($header).'<br>';
if (!empty($_POST) && $_POST['export-woo-orders'] == '1') {
	$spreadsheet = new Spreadsheet();
	$sheet = $spreadsheet->getActiveSheet();

	$categories = $_POST['categories'];
	// print_r($categories);
	$cat_data = array();
	$total_cells_needed = 5 + 3;
	$pre_headers_values = array('*', '*', '*', '*', '*');
	$headTexts = array();
	$product_ids = [];
	$headers_values = array('DATE', 'TIME', 'CLIENT NAME', 'PHONE', 'SHOP');
	$filterProducts = array();
	$filterCategories = array();

	// foreach ($categories as $k => $category) {
	// 	$tempCount = 0;
	// 	$pargs = array(
	// 		'product_category_id' 	=> $category,
	// 		'status' 				=> 'publish',
	// 		'limit' 				=> -1,
	// 		'orderby' 				=> 'title',
	// 		'order' 				=> 'ASC',
	// 	);
	// 	$_products = wc_get_products($pargs);
	// 	foreach ($_products as $i => $a_product) {
	// 		if (!empty(get_orders_ids_by_product_id($a_product->get_id()))) {
	// 			$tempCount++;
	// 		}
	// 	}
	// 	if ($tempCount > 0) {
	// 		$filterCategories[] = $category;
	// 	}
	// }
	// foreach ($filterCategories as $k => $category) {
	// 	$cat_obj = get_term_by('id', $category, 'product_cat');
	// 	$pargs = array(
	// 		'product_category_id' 	=> $category,
	// 		'status' 				=> 'publish',
	// 		'limit' 				=> -1,
	// 		'orderby' 				=> 'title',
	// 		'order' 				=> 'ASC',
	// 	);
	// 	$f_products = wc_get_products($pargs);
	// 	foreach ($f_products as $i => $fproduct) {
	// 		if (!empty(get_orders_ids_by_product_id($fproduct->get_id()))) {
	// 			$filterProducts[] = $fproduct;
	// 		}
	// 	}

	// 	// echo count($filterProducts) . "<br>";
	// 	// exit();
	// 	$total_cells_needed += count($filterProducts);
	// 	$cat_data[$cat_obj->name] = ['total' => count($filterProducts)];
	// 	$pre_headers_values[] = $cat_obj->name;
	// 	$headTexts[$k] = $cat_obj->name;
	// 	foreach ($filterProducts as $i => $f_product) {
	// 		// echo $f_product->get_id() . '-' . $f_product->get_title();
	// 		// echo "<br>";
	// 		$cat_data[$cat_obj->name]['products'][] = ['id' => $f_product->get_id(), 'title' => $f_product->get_title()];
	// 		$headers_values[] = $f_product->get_title();
	// 		$product_ids[] = $f_product->get_id();
	// 		if ($i != 0) {
	// 			$pre_headers_values[] = $k . '-';
	// 		}
	// 	}
	// 	$filterProducts = [];
	// }
	foreach ($categories as $k => $category) {
		$cat_obj = get_term_by('id', $category, 'product_cat');
		$pargs = array(
			'product_category_id' 	=> $category,
			'status' 				=> 'publish',
			'limit' 				=> -1,
			'orderby' 				=> 'title',
			'order' 				=> 'ASC',
		);
		$products = wc_get_products($pargs);
		// print_r($products);
		// echo count($products) . "<br>";
		$total_cells_needed += count($products);
		$cat_data[$cat_obj->name] = ['total' => count($products)];
		$pre_headers_values[] = $cat_obj->name;
		$headTexts[$k] = $cat_obj->name;
		foreach ($products as $i => $product) {
			// echo $product->get_id() . '-' . $product->get_title();
			// echo "<br>";
			$cat_data[$cat_obj->name]['products'][] = ['id' => $product->get_id(), 'title' => $product->get_title()];
			$headers_values[] = $product->get_title();
			$product_ids[] = $product->get_id();
			if ($i != 0) {
				$pre_headers_values[] = $k . '-';
			}
		}
	}
	$pre_headers_values[] = '+'; // REMARKS
	$pre_headers_values[] = '+'; // ORDER NUMBER
	$pre_headers_values[] = '+'; // TOTAL AMOUNT

	$headers_values[] = 'REMARKS';
	$headers_values[] = 'ORDER NUMBER';
	$headers_values[] = 'TOTAL AMOUNT';
	// echo $total_cells_needed;
	// echo "<pre>";
	// print_r($cat_data);
	// echo "</pre>";
	$letters = [];

	$totalSets = ceil($total_cells_needed / count(range('A', 'Z')));
	$wrapArray = range('A', 'Z');
	foreach (range('A', 'Z') as $char) {
		if (count($letters) < $total_cells_needed) {
			$letters[] = $char;
		}
	}
	if ($totalSets > 1) {
		foreach ($wrapArray as $val) {
			foreach (range('A', 'Z') as $char) {
				if (count($letters) < $total_cells_needed) {
					$letters[] = $val . $char;
				}
			}
		}
	}
	$last_3_array = array_slice($letters, -3, 3, false);
	$products_letters = array_slice($letters, 5, (count($letters) - 8), false);
	// echo "<pre>";
	// print_r($pre_headers_values);
	// print_r($letters);
	// print_r($headers_values);
	// print_r($cat_data);
	// echo count($letters); // 26
	// print_r($last_3_array);
	// print_r($filterProducts);
	// print_r($headTexts);
	// print_r($product_ids);
	// echo "</pre>";
	// exit;
	// var_dump(array_key_exists('Pasteles Otoño', $cat_data));
	// var_dump($cat_data['Pasteles Otoño']['total']);
	// echo array_search('Pasteles Otoño', $pre_headers_values);
	$headIndexes = array();
	foreach ($headTexts as $headText) {
		$headIndexes[] = array_search($headText, $pre_headers_values);
	}
	// print_r($headIndexes);
	$mergeCells = array('A1:E1');
	// exit;
	// PRE HEADER
	foreach ($pre_headers_values as $key => $value) {
		$filterVal = str_replace('-', '', str_replace('+', '', str_replace('*', '', $value)));

		foreach ($filterCategories as $i => $cat) {
			$filterVal = str_replace($i, '', $filterVal);
		}
		$sheet->setCellValue($letters[$key] . '1', $filterVal);
	}
	foreach ($headTexts as $k => $headText) {
		$Idx = $headIndexes[$k];
		$total = $cat_data[$headText]['total'];
		if ($total > 1) {
			$mergeCells[] = $letters[$Idx] . '1:' . $letters[($Idx + $total) - 1] . '1';
		} else {
			$mergeCells[] = $letters[$Idx] .  '1';
		}
	}
	// print_r($mergeCells);
	// exit;

	$sheet->getRowDimension('1')->setRowHeight(50); //set height of 1st row
	foreach ($mergeCells as $k => $finalCell) {
		if (strlen($finalCell) > 2) {
			$sheet->mergeCells($finalCell);
		}

		if ($k != '0') {
			$sheet->getStyle($finalCell)->getFill()->applyFromArray(
				[
					'fillType' => 'solid',
					'startColor' => [
						'rgb' => 'EBC852'
					],
					'endColor' => [
						'rgb' => 'FFFFFFFF'
					],
				]
			);

			$sheet->getStyle($finalCell)->getAlignment()->applyFromArray(
				[
					'horizontal'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
					'vertical'     => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
					'textRotation' => 0,
					'wrapText'     => TRUE
				]
			);

			$sheet->getStyle($finalCell)->getFont()->applyFromArray(
				[
					'name' => 'Calibri',
					'bold' => TRUE,
					'italic' => FALSE,
					'underline' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE,
					'strikethrough' => FALSE,
					'color' => [
						'rgb' => '000000'
					]
				]
			);
			$sheet->getStyle($finalCell)->getBorders()->getRight()->applyFromArray(
				[
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
					'color' => [
						'rgb' => '000000'
					]
				]
			);
		}
	}
	//Headers
	foreach ($letters as $k => $letter) {
		$sheet->setCellValue($letter . '2', $headers_values[$k]);
		$sheet->getStyle($letter . '2')->getAlignment()->applyFromArray(
			[
				'horizontal'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
				'vertical'     => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
				'textRotation' => 0,
				'wrapText'     => TRUE
			]
		);

		$sheet->getStyle($letter . '2')->getFont()->applyFromArray(
			[
				'name' => 'Calibri',
				'bold' => TRUE,
				'italic' => FALSE,
				'underline' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE,
				'strikethrough' => FALSE,
				'color' => [
					'rgb' => '000000'
				]
			]
		);
		$sheet->getStyle($letter . '2')->getBorders()->getRight()->applyFromArray(
			[
				'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				'color' => [
					'rgb' => '000000'
				]
			]
		);
	}
	$sheet->getRowDimension('2')->setRowHeight(25); //set height of 2nd row

	$args =  array(
		'limit' 	=> -1,
		'orderby' 	=> 'date',
		'order' 	=> 'DESC',
		'return' 	=> 'ids',
		'status'	=> 'processing',
	);

	$final_order = array();

	//print_r($args);
	$query = new WC_Order_Query($args);
	$final_order = $query->get_orders();
	// echo "<pre>";
	// print_r($final_order);
	// echo "</pre>";
	// exit;
	$rowCount = 3;
	$j = 5;

	foreach ($final_order as $key => $order_id) {
		$t = 0;
		foreach ($products_letters as $i => $products_letter) {
			$t += calculate_products_count($order_id, $product_ids[$i]);
		}
		if ($t == 0) continue;

		$order = wc_get_order($order_id);
		$items = $order->get_items();

		$delivery_date = $order->get_meta('pi_delivery_date', true);
		$sheet->setCellValue('A' . $rowCount, str_replace('/', '-', $delivery_date));

		$delivery_time = $order->get_meta('pi_delivery_time', true);
		$sheet->setCellValue('B' . $rowCount, $delivery_time);

		$sheet->setCellValue('C' . $rowCount, $order->get_billing_first_name() . ' ' . $order->get_billing_last_name());

		$sheet->setCellValue('D' . $rowCount, $order->get_billing_phone());

		$sheet->setCellValue('E' . $rowCount, ''); //SHOP

		foreach ($products_letters as $i => $products_letter) {
			$sheet->setCellValue($products_letter . $rowCount,  calculate_products_count($order_id, $product_ids[$i]));
			$sheet->getStyle($products_letter . $rowCount)->getAlignment()->applyFromArray(
				[
					'horizontal'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
					'vertical'     => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
					'textRotation' => 0,
					'wrapText'     => TRUE
				]
			);
		}

		$sheet->setCellValue($last_3_array[0] . $rowCount, ''); // REMARKS
		$sheet->setCellValue($last_3_array[1] . $rowCount, $order->get_order_number()); // ORDER NUMBER
		$sheet->setCellValue($last_3_array[2] . $rowCount, number_format($order->get_total(), 2, '.')); // TOTAL AMOUNT
		$sheet->getStyle($last_3_array[2] . $rowCount)->getAlignment()->applyFromArray(
			[
				'horizontal'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
				'vertical'     => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
				'textRotation' => 0,
				'wrapText'     => TRUE
			]
		);
		$rowCount++;
	}

	$sheet->setCellValue($last_3_array[2] . $rowCount, '=SUM(' . $last_3_array[2] . '3:' . $last_3_array[2] . ($rowCount - 1) . ')');
	$sheet->getStyle($last_3_array[2] . $rowCount)->getAlignment()->applyFromArray(
		[
			'horizontal'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
			'vertical'     => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
			'textRotation' => 0,
			'wrapText'     => TRUE
		]
	);
	$sheet->getStyle($last_3_array[2] . $rowCount)->getFont()->applyFromArray(
		[
			'name' => 'Calibri',
			'bold' => TRUE,
			'italic' => FALSE,
			'underline' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE,
			'strikethrough' => FALSE,
			'color' => [
				'rgb' => '000000'
			]
		]
	);

	foreach ($products_letters as $i => $products_letter) {
		$sheet->setCellValue($products_letter . $rowCount, '=SUM(' . $products_letter . '3:' . $products_letter . ($rowCount - 1) . ')');
		$sheet->getStyle($products_letter . $rowCount)->getAlignment()->applyFromArray(
			[
				'horizontal'   => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
				'vertical'     => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
				'textRotation' => 0,
				'wrapText'     => TRUE
			]
		);

		$sheet->getStyle($products_letter . $rowCount)->getFont()->applyFromArray(
			[
				'name' => 'Calibri',
				'bold' => TRUE,
				'italic' => FALSE,
				'underline' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE,
				'strikethrough' => FALSE,
				'color' => [
					'rgb' => '000000'
				]
			]
		);
	}
	// $sheet->setCellValue('F' . $rowCount,  '=SUM(F3:F' . ($rowCount - 1) . ')');

	// foreach ($products_letters as $i => $products_letter) {
	// 	$sheet->setCellValue($products_letter . $rowCount,  calculate_products_count($order_id, $product_ids[$i]));
	// }

	// $product_cell_array = array('A', 'B', 'C', 'D', 'E', $last_3_array[0], $last_3_array[1], $last_3_array[2]);

	// print_r($product_cell_array);

	// $result = array_diff($letters, $product_cell_array);
	// print_r($result);
	// exit;
	$filename = "orders-" . date("Y-m-d-H-i-s") . ".xlsx";
	$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
	$fileName = $fileName . '.xlsx';
	ob_end_clean();
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="' . $fileName . '"');
	header('Cache-Control: max-age=0');
	header("Content-Disposition: attachment; filename=" . $filename);
	exit($writer->save('php://output'));
}

function calculate_products_count($order_id, $product_id)
{
	$order = wc_get_order($order_id);
	$items = $order->get_items();
	// echo count($items) . '<br>';
	// print_r($items);
	$total_products = 0;
	if (count($items) > 0) {
		foreach ($items as $item_id => $item) {
			if ($item->get_product_id() == $product_id) {
				$total_products +=  $item->get_quantity();
			}
		}
	}
	return $total_products;
}
// $order = wc_get_order(16537);
// var_dump($order);
// $items = $order->get_items();
// var_dump($items);
// echo count($items) . '<br>';
// foreach ($items as $item_id => $item) {
// 	print_r($item) . '<br>';
// }
// var_dump(check_if_atleast_one_order_exists(15402));

function get_orders_ids_by_product_id($product_id)
{

	global $wpdb;
	$order_status = ['wc-processing'];
	$prepared_statement = "SELECT order_items.order_id FROM {$wpdb->prefix}woocommerce_order_items as order_items LEFT JOIN {$wpdb->prefix}woocommerce_order_itemmeta as order_item_meta ON order_items.order_item_id = order_item_meta.order_item_id JOIN cjlq_wc_orders AS posts ON order_items.order_id = posts.id  WHERE posts.status ='wc-processing' AND order_items.order_item_type = 'line_item' AND order_item_meta.meta_key = '_product_id' AND order_item_meta.meta_value = '" . $product_id . "' ORDER BY order_items.order_id DESC";
	$results = $wpdb->get_col($prepared_statement);

	return $results;
}
// print_r(get_orders_ids_by_product_id(15402));
// print_r(get_orders_ids_by_product_id(36036));
?>
<style>
	label.wce_label {
		display: inline-block;
	}

	#order_data {
		padding: 23px 24px 12px;
	}

	select#wc_categories {
		max-width: 100%;
	}

	.select2-container {
		width: 50% !important;
	}

	li.select2-selection__choice {
		height: 20px;
		padding-top: 10px !important;
		padding-bottom: 5px !important;
		margin-bottom: 3px !important;
	}

	span.select2-selection.select2-selection--multiple {
		font-size: 14px;
	}

	.select2-container .select2-search--inline {
		font-size: 16px;
		padding: 5px;
	}

	input.select2-search__field {
		padding-left: 5px !important;
	}

	span.select2-selection__clear {
		font-size: 18px;
		color: #d63638;
	}
</style>

<link rel='stylesheet' id='jquery-ui-style-css' href='//ajax.googleapis.com/ajax/libs/jqueryui/1.11.3/themes/smoothness/jquery-ui.css' type='text/css' media='all' />
<script src="https://lapastisseriabarcelona.com/wp-content/plugins/woocommerce/assets/js/select2/select2.min.js"></script>
<div class="wrap automatewoo-page">
	<h1 class="wp-heading-inline"><?php esc_html_e('WOO Order Export', 'wc-order-export');  ?></h1>
	<hr class="wp-header-end">
	<div class="postbox">
		<?php
		$args = array(
			'hide_empty' => true,
			'exclude' => '298',
			'parent'   => 0,
		);
		$terms = get_terms('product_cat', $args);
		// print_r($terms);
		?>
		<div id="order_data">
			<form method="post" class="form-wrap">
				<input type="hidden" name="export-woo-orders" value="1" />

				<div class="form-field form-required term-name-wrap inited inited_media_selector">
					<h3 for="wc_categories">
						<?php esc_html_e('Select', 'woocommerce');  ?>
						<?php esc_html_e('Category', 'woocommerce');  ?></h3>
					<select name="categories[]" id="wc_categories" class="wc_categories" required>
						<?php
						if ($terms) {
							foreach ($terms as $term) {
								echo '<option value="' . $term->term_id . '">' . $term->name . '</option>';
							}
						}
						?>
					</select>
				</div>

				<br><br>
				<button type="submit" class="button button-primary"><?php esc_html_e('Export', 'woocommerce') ?></button>
			</form>
		</div>
	</div>
</div>
<script>
	jQuery(function($) {
		jQuery('select#wc_categories').select2({
			placeholder: {
				id: '-1', // the value of the option
				text: "<?php esc_html_e('Select an option', 'woocommerce') ?>",
			},
			allowClear: true,
			multiple: true,
			language: {
				noResults: function(params) {
					return "<?php esc_html_e('No results found', 'woocommerce') ?>";
				}
			},
		});
		jQuery('select#wc_categories').val('').trigger('change');
	});
</script>