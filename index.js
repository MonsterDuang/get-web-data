// 你可以在这里添加更多的代码来操作宿主网页的window对象

$(document).ready(() => {
	let tabList = [];
	let activeTab = null;
	let smuId = '';
	let modelId = '';
	let productList = [];

	// 动态添加样式
	const style = document.createElement('style');
	style.innerHTML = `
				::-webkit-scrollbar {
          width: 6px;
          height: 6px;
        }
        ::-webkit-scrollbar-track {
          background-color: transparent;
          border-radius: 3px;
        }
        ::-webkit-scrollbar-thumb {
          background-color: rgba(0, 123, 255, 0.5);
          border-radius: 3px;
        }
				#moveButton {
					position: fixed;
					right: 20px;
					top: 50%;
					width: 50px;
					height: 50px;
					background-color: #007bff;
					color: #fff;
					border-radius: 50%;
					display: flex;
					align-items: center;
					justify-content: center;
					cursor: pointer;
					z-index: 1000;
					user-select: none;
				}
        #tabs {
          display: flex;
          justify-content: flex-start;
          margin-top: 20px;
					overflow-x: auto;
        }
        .tab {
          padding: 10px 20px;
          margin: 0 10px;
          cursor: pointer;
					flex: none;
          border-bottom: 2px solid transparent;
        }
        .tab.active {
          border-bottom: 2px solid rgb(0, 123, 255);
        }
        #tableContent {
          margin-top: 20px;
          background-color: #f9f9f9;
          border-radius: 10px;
          box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
          flex: 1;
          overflow-y: auto;
					border: 1px solid #ddd;
        }
        table {
          width: 100%;
          border-collapse: collapse;
        }
        th, td {
          padding: 10px;
          border-bottom: 1px solid #ddd;
          text-align: left;
        }
        td a {
          color: #007bff;
          text-decoration: none;
        }
        td a:hover {
          text-decoration: underline;
        }
        th {
          background-color: #f2f2f2;
          position: sticky;
          top: 0;
          z-index: 1;
        }
        .product-image {
          cursor: pointer;
          width: 50px;
          height: 50px;
        }
        .bottomBtn {
          text-align: right;
          margin-top: 20px;
        }
        #exportButton {
          padding: 8px 16px;
          background-color: #007bff;
          color: #fff;
          border: none;
          border-radius: 5px;
          cursor: pointer;
        }
        #exportButton:hover {
          opacity: 0.8;
        }
				#imageModal {
					position: fixed;
					top: 0;
					left: 0;
					width: 100%;
					height: 100%;
					background-color: rgba(0, 0, 0, 0.8);
					display: flex;
					align-items: center;
					justify-content: center;
					z-index: 10000;
				}
				#imageModalContent {
					position: relative;
					padding: 20px;
					background-color: #fff;
					border-radius: 10px;
					box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
				}
				#imageModalClose {
					position: absolute;
					top: 10px;
					right: 10px;
					font-size: 24px;
					cursor: pointer;
					color: #000;
				}
				#imageModalImg {
					max-width: 100%;
					max-height: 80vh;
				}
				#popup {
					position: fixed;
					top: 0;
					left: 0;
					width: 100%;
					height: 100%;
					padding: 100px 200px;
					box-sizing: border-box;
					z-index: 9999;
					display: flex;
					align-items: center;
					justify-content: center;
					background-color: rgba(0, 0, 0, 0.5);
				}
				#popupContent {
					position: relative;
					width: 100%;
					height: 100%;
					padding: 20px;
					box-sizing: border-box;
					background-color: #fff;
					border-radius: 10px;
					box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
					display: flex;
					flex-direction: column;
				}
				#popupClose {
					position: absolute;
					top: 10px;
					right: 10px;
					font-size: 24px;
					cursor: pointer;
				}
				#popupTitle {
					margin: 0;
					padding: 0;
					text-align: center;
					font-size: 24px;
				}
      `;
	document.head.appendChild(style);

	// 导出表格数据为Excel
	const exportExcel = async () => {
		const workbook = new ExcelJS.Workbook();
		const worksheet = workbook.addWorksheet('产品数据');

		const title = $('#popupTitle').text();
		const label = activeTab.label;

		// 添加标题行（在第一行插入）
		const titleRow = worksheet.addRow([`类目：${title}【${label}】`]); // 插入到第一行
		// 合并单元格（A1到C1）
		worksheet.mergeCells('A1:G1'); // 或使用数字参数：mergeCells(1, 1, 1, 3)
		// 设置标题样式
		const titleStyle = {
			font: {
				name: '微软雅黑',
				bold: true,
				size: 16,
				color: { argb: 'FFFFFFFF' }, // 白色字体
			},
			fill: {
				type: 'pattern',
				pattern: 'solid',
				fgColor: { argb: 'FF4F81BD' }, // 深蓝色背景
			},
			alignment: {
				vertical: 'middle',
				horizontal: 'center',
			},
		};
		// 应用标题样式（只需要设置合并区域的第一个单元格）
		titleRow.getCell(1).style = titleStyle;

		// 设置标题行高度
		titleRow.height = 40;
		// 设置表头
		const headers = ['序号', '产品图片', '产品描述', '销售价格', '订单数量', '最低起订量', '详情链接'];
		// 添加表头行（现在变为第二行）
		const headerRow = worksheet.addRow(headers);
		// 设置表头样式（复用之前的样式配置）
		const headerStyle = {
			font: {
				name: '微软雅黑',
				bold: true,
				size: 12,
				color: { argb: 'FF000000' }, // 黑色字体
			},
			fill: {
				type: 'pattern',
				pattern: 'solid',
				fgColor: { argb: 'FFD3D3D3' }, // 浅灰色背景
			},
			border: {
				top: { style: 'thin', color: { argb: 'FF000000' } },
				bottom: { style: 'thin', color: { argb: 'FF000000' } },
				left: { style: 'thin', color: { argb: 'FF000000' } },
				right: { style: 'thin', color: { argb: 'FF000000' } },
			},
			alignment: {
				vertical: 'middle',
				horizontal: 'center',
			},
		};
		headerRow.eachCell((cell) => (cell.style = headerStyle));
		headerRow.height = 25;

		// 添加数据
		productList.forEach(async (product, index) => {
			const imageUrl = product.image.split('_')[0];
			const row = worksheet.addRow([
				index + 1,
				{ text: '查看', hyperlink: imageUrl },
				product.subject,
				product.price,
				smuId === '6622646540' ? product.orders : smuId === '1326427270' ? product.salesCount : smuId === '6809349780' ? product.prodSold180 : '',
				product.moq,
				{ text: '跳转', hyperlink: `https:${product.detail}` },
			]);

			const rowIdx = index + 3;
			worksheet.getRow(rowIdx).height = 20;

			// 遍历单元格设置超链接样式
			row.eachCell((cell) => {
				if (cell.value && cell.value.hyperlink) {
					// 超链接专属样式
					cell.style = {
						font: {
							color: { argb: 'FF0000FF' }, // 蓝色
							underline: true,
							name: 'Calibri',
						},
						alignment: {
							vertical: 'middle',
							horizontal: 'center',
						},
					};
					// 可选：移除默认的紫色下划线（覆盖Excel默认样式）
					cell.value = {
						text: cell.value.text,
						hyperlink: cell.value.hyperlink,
					};
				} else {
					// 普通单元格样式
					cell.style = {
						font: { name: 'Calibri' },
						alignment: { vertical: 'middle', horizontal: 'center' },
					};
				}
			});
		});

		// 自动调整列宽（按内容）
		worksheet.columns.forEach((column) => {
			let maxLength = 0;
			column.eachCell({ includeEmpty: true }, (cell) => {
				const columnLength = cell.value ? cell.value.toString().length : 10;
				if (columnLength > maxLength) maxLength = columnLength;
			});
			column.width = maxLength * 2;
		});

		// 导出Excel文件
		const buffer = await workbook.xlsx.writeBuffer();
		const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
		const url = URL.createObjectURL(blob);
		const a = document.createElement('a');
		a.href = url;
		a.download = `${title}-${label}-数据列表.xlsx`;
		a.click();
		URL.revokeObjectURL(url);
	};

	// 请求对应的数据
	const getProductList = (item) => {
		const api = 'https://insights.alibaba.com/openservice/gatewayService?language=zh&pageNo=1&pageSize=500&';
		const params =
			smuId === '6622646540'
				? `modelId=${modelId}&cardId=${item.cardId}&cardType=${item.cardType}`
				: smuId === '1326427270'
				? `modelId=${modelId}&categoryIds=${item.sortIds}`
				: smuId === '6809349780'
				? `modelId=${modelId}&categoryIds=${item.categoryId || ''}&newsType=101002913`
				: '';
		$.ajax({
			url: `${api}${params}`,
			method: 'GET',
			dataType: 'json',
			success: ({ data }) => {
				productList =
					smuId === '6622646540'
						? data.list[0].productList
						: smuId === '1326427270'
						? data.list.map((x) => x.data)
						: smuId === '6809349780'
						? data.list.map((x) => x.product)
						: [];
				if (!$('#exportButton').length) $('#popupContent').append('<div class="bottomBtn"><button id="exportButton">导出Excel</button></div>');
				const $table = $('<table></table>');
				const $thead = $(`
              <thead>
                <tr>
                  <th>序号</th>
                  <th>产品图片</th>
                  <th>产品描述</th>
                  <th>销售价格</th>
                  <th>订单数量</th>
                  <th>最低起订量</th>
                  <th>详情链接</th>
                </tr>
              </thead>
            `);
				const $tbody = $('<tbody></tbody>');

				productList.forEach((product, index) => {
					if (!product) return;
					const $row = $(`
                <tr>
                  <td>${index + 1}</td>
                  <td><img src="${product.image}" class="product-image" /></td>
                  <td>${product.subject}</td>
                  <td>${product.price}</td>
                  <td>${
										smuId === '6622646540'
											? product.orders
											: smuId === '1326427270'
											? product.salesCount
											: smuId === '6809349780'
											? product.prodSold180
											: ''
									}</td>
                  <td>${product.moq}</td>
                  <td><a href="${product.detail}" target="_blank">跳转</a></td>
                </tr>
              `);
					$tbody.append($row);
				});

				$table.append($thead).append($tbody);
				$('#tableContent').html($table);

				$('#exportButton').on('click', exportExcel);

				// 添加图片点击放大功能
				$('.product-image').on('click', function () {
					const src = $(this).attr('src');
					const $imageModal = $(`
                <div id="imageModal">
                  <div id="imageModalContent">
                    <span id="imageModalClose">&times;</span>
                    <img src="${src.split('_')[0]}" id="imageModalImg" />
                  </div>
                </div>
              `);
					$('body').append($imageModal);

					// 关闭图片模态框
					$('#imageModalClose').on('click', () => {
						$imageModal.remove();
					});
				});
			},
			error: (error) => {
				console.error('Error fetching data:', error);
				$('#tableContent').html('<p>Failed to fetch data.</p>');
			},
		});
	};

	// 点击按钮时创建弹出框
	const openDialog = () => {
		const $popup = $(`
        <div id="popup">
          <div id="popupContent">
            <span id="popupClose">&times;</span>
            <h1 id="popupTitle">${
							smuId === '6622646540' ? tabList[0].rankWord : smuId === '1326427270' ? '新品上架' : smuId === '6809349780' ? '热门促销' : ''
						}</h1>
            <div id="tabs"></div>
            <div id="tableContent"></div>
          </div>
        </div>
      `);
		$('body').append($popup);

		// 生成tabs
		const $tabs = $('#tabs');
		tabList.forEach((item, index) => {
			const $tab = $(`<div class="tab">${item.label}</div>`);
			$tabs.append($tab);
			$tab.on('click', () => {
				if ($tab.hasClass('active')) return;
				$('.tab').removeClass('active');
				$('.tab').attr('id', '');
				activeTab = item;
				$tab.addClass('active');
				$tab.attr('id', 'active');
				document.getElementById('active').scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
				$('#tableContent').empty();
				$('.bottomBtn').remove();
				getProductList(item);
			});
			if (index === 0) {
				$tab.addClass('active');
				activeTab = item;
				getProductList(item);
			}
		});

		// 关闭弹出框
		$('#popupClose').on('click', () => {
			$popup.remove();
		});
	};

	const createDom = () => {
		// 创建圆形按钮
		const $button = $('<div id="moveButton">导出</div>');
		$('body').append($button);
		$button.hover(
			() => {
				$button.css('opacity', 0.8);
			},
			() => {
				$button.css('opacity', 1);
			}
		);

		// 添加拖动功能
		let isDragging = false;
		let startY, startTop;

		$button.on('mousedown', (e) => {
			isDragging = true;
			startY = e.clientY;
			startTop = parseInt($button.css('top'), 10);
			$('body').on('mousemove', onMouseMove);
			$('body').on('mouseup', onMouseUp);
		});

		const onMouseMove = (e) => {
			if (isDragging) {
				const dy = e.clientY - startY;
				$button.css('top', startTop + dy + 'px');
			}
		};

		const onMouseUp = () => {
			isDragging = false;
			$('body').off('mousemove', onMouseMove);
			$('body').off('mouseup', onMouseUp);
		};

		$button.on('click', openDialog);
	};

	const getList = () => {
		const scriptArr = document.querySelectorAll('script');
		for (let item of scriptArr) {
			smuId = item.dataset.suspenseModuleUuid;
			if (['6622646540', '1326427270', '6809349780'].includes(smuId)) {
				let dataStr = item.innerHTML.split(' = ')[1].split('window.moduleHtml_')[0];
				// 去除多余的字符，确保JSON格式正确
				dataStr = dataStr.trim().replace(/;$/, '');
				try {
					const dataObj = JSON.parse(dataStr);
					if (smuId === '6622646540') {
						// 热门爆品
						const { _serverData, params } = dataObj._fdl_request.requestList[0];
						tabList = _serverData.list;
						modelId = JSON.parse(params).modelId;
					} else if (smuId === '1326427270') {
						// 新品上架
						const {
							_fdl_request_tab: { _serverData },
							_fdl_request_waterfall: { params },
						} = dataObj._fdl_request.requestCommon;
						tabList = _serverData.tabList;
						modelId = JSON.parse(params).modelId;
					} else if (smuId === '6809349780') {
						// 热门促销
						const {
							_fdl_request_nav: { _serverData },
							_fdl_request_theme_waterfall: { params },
						} = dataObj._fdl_request.requestCommon;
						tabList = _serverData.list;
						modelId = JSON.parse(params).modelId;
					}
					createDom();
				} catch (error) {
					console.error('Error parsing JSON:', error);
				}
			}
		}
	};

	getList();
});
