// 你可以在这里添加更多的代码来操作宿主网页的window对象

$(document).ready(() => {
	let list = [];
	let productList = [];
	let modelId = '';
	let activeTab = null;

	const createDom = () => {
		// 创建圆形按钮
		const $button = $('<div id="moveButton">导出</div>');
		$('body').append($button);

		// 设置按钮样式
		$button.css({
			position: 'fixed',
			right: '20px',
			top: '50%',
			width: '50px',
			height: '50px',
			backgroundColor: '#007bff',
			color: '#fff',
			borderRadius: '50%',
			display: 'flex',
			alignItems: 'center',
			justifyContent: 'center',
			cursor: 'pointer',
			zIndex: 1000,
			userSelect: 'none',
		});
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

		// 点击按钮时创建弹出框
		$button.on('click', () => {
			const $popup = $(`
        <div id="popup">
          <div id="popupContent">
            <span id="popupClose">&times;</span>
            <h1 id="popupTitle"></h1>
            <div id="tabs"></div>
            <div id="tabContent"></div>
            <div class="bottomBtn"><button id="exportButton">导出Excel</button></div>
          </div>
        </div>
      `);
			$('body').append($popup);

			// 设置弹出框样式
			$popup.css({
				position: 'fixed',
				top: 0,
				left: 0,
				width: '100%',
				height: '100%',
				padding: '30px',
				zIndex: 9999,
				display: 'flex',
				alignItems: 'center',
				justifyContent: 'center',
				backgroundColor: 'rgba(0, 0, 0, 0.5)',
			});

			$('#popupContent').css({
				position: 'relative',
				width: '100%',
				height: '100%',
				padding: '20px',
				boxSizing: 'border-box',
				backgroundColor: '#fff',
				borderRadius: '10px',
				boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)',
				display: 'flex',
				flexDirection: 'column',
			});

			$('#popupClose').css({
				position: 'absolute',
				top: '10px',
				right: '10px',
				fontSize: '24px',
				cursor: 'pointer',
			});

			$('#popupTitle').text(list[0].rankWord).css({
				margin: '0',
				padding: '0',
				textAlign: 'center',
				fontSize: '24px',
			});

			// 检查图片格式的示例函数
			const isValidImageType = (buffer) => {
				const header = buffer.slice(0, 4).toString('hex');
				return (
					header.startsWith('89504e47') || // PNG
					header.startsWith('ffd8ffe0') || // JPEG
					header.startsWith('47494638') // GIF
				);
			};

			// JPG转PNG核心函数
			const convertJPGtoPNGBuffer = async (imageUrl) => {
				// 获取图片 Blob（可选，如果图片有跨域问题，请确保服务器允许）
				const response = await fetch(imageUrl);
				const blob = await response.blob();

				return new Promise((resolve, reject) => {
					const img = new Image();
					// 处理跨域问题（需要服务器允许跨域访问）
					img.crossOrigin = 'Anonymous';
					img.onload = async () => {
						const canvas = document.createElement('canvas');
						canvas.width = img.width;
						canvas.height = img.height;
						const ctx = canvas.getContext('2d');
						ctx.drawImage(img, 0, 0);

						// 转换为 PNG 格式
						canvas.toBlob(async (pngBlob) => {
							if (!pngBlob) {
								return reject(new Error('转换失败'));
							}
							try {
								// 获取 buffer
								const buffer = await pngBlob.arrayBuffer();
								resolve(buffer);
							} catch (error) {
								reject(error);
							}
						}, 'image/png');
					};
					img.onerror = reject;
					// 使用 URL.createObjectURL 为 blob 创建临时 URL
					img.src = URL.createObjectURL(blob);
				});
			};

			// 请求对应的数据
			const getProductList = (item) => {
				$.ajax({
					url: `https://insights.alibaba.com/openservice/gatewayService?modelId=${modelId}&cardId=${item.cardId}&cardType=${item.cardType}&language=zh`,
					method: 'GET',
					dataType: 'json',
					success: ({ data }) => {
						productList = data.list[0].productList;
						const $table = $('<table></table>');
						const $thead = $(`
              <thead>
                <tr>
                  <th>产品图片</th>
                  <th>产品描述</th>
                  <th>销售价格</th>
                  <th>订单数量</th>
                  <th>观看次数</th>
                  <th>人气度</th>
                  <th>最低起订量</th>
                  <th>详情链接</th>
                </tr>
              </thead>
            `);
						const $tbody = $('<tbody></tbody>');

						productList.forEach((product) => {
							const $row = $(`
                <tr>
                  <td><img src="${product.image}" class="product-image" /></td>
                  <td>${product.subject}</td>
                  <td>${product.price}</td>
                  <td>${product.orders}</td>
                  <td>${product.views}</td>
                  <td>${product.rankScore}</td>
                  <td>${product.moq}</td>
                  <td><a href="${product.detail}" target="_blank">跳转</a></td>
                </tr>
              `);
							$tbody.append($row);
						});

						$table.append($thead).append($tbody);
						$('#tabContent').html($table);

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

							// 设置图片模态框样式
							$imageModal.css({
								position: 'fixed',
								top: 0,
								left: 0,
								width: '100%',
								height: '100%',
								backgroundColor: 'rgba(0, 0, 0, 0.8)',
								display: 'flex',
								alignItems: 'center',
								justifyContent: 'center',
								zIndex: 10000,
							});

							$('#imageModalContent').css({
								position: 'relative',
								padding: '20px',
								backgroundColor: '#fff',
								borderRadius: '10px',
								boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)',
							});

							$('#imageModalClose').css({
								position: 'absolute',
								top: '10px',
								right: '10px',
								fontSize: '24px',
								cursor: 'pointer',
								color: '#000',
							});

							$('#imageModalImg').css({
								maxWidth: '100%',
								maxHeight: '80vh',
							});

							// 关闭图片模态框
							$('#imageModalClose').on('click', () => {
								$imageModal.remove();
							});
						});
					},
					error: (error) => {
						console.error('Error fetching data:', error);
						$('#tabContent').html('<p>Failed to fetch data.</p>');
					},
				});
			};

			// 导出表格数据为Excel
			$('#exportButton').on('click', async () => {
				const workbook = new ExcelJS.Workbook();
				const worksheet = workbook.addWorksheet('产品数据');

				const title = $('#popupTitle').text();
				const label = activeTab.label;

				// 添加标题行（在第一行插入）
				const titleRow = worksheet.addRow([`类目：${title}【${label}】`]); // 插入到第一行
				// 合并单元格（A1到C1）
				worksheet.mergeCells('A1:H1'); // 或使用数字参数：mergeCells(1, 1, 1, 3)
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
				const headers = ['产品图片', '产品描述', '销售价格', '订单数量', '观看次数', '人气度', '最低起订量', '详情链接'];
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
					const row = worksheet.addRow([
						'',
						product.subject,
						product.price,
						product.orders,
						product.views,
						product.rankScore,
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

					// 处理图片
					const imageUrl = product.image.split('_')[0];
					// 获取图片的buffer数据
					const pngBuffer = await convertJPGtoPNGBuffer(imageUrl);
					const PngBuffer = new Uint8Array(pngBuffer);
					console.log('PngBuffer::: ', PngBuffer);

					const imageId = workbook.addImage({
						buffer: PngBuffer,
						extension: 'png', // 根据实际类型设置
					});

					// 验证图片格式
					if (!isValidImageType(PngBuffer)) {
						throw new Error('不支持的图片格式');
					}

					worksheet.addImage(imageId, {
						tl: { col: 0, row: rowIdx - 1 },
						ext: { width: 10, height: 10 },
						editAs: 'absolute', // 绝对定位
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
			});

			// 生成tabs
			const $tabs = $('#tabs');
			list.forEach((item, index) => {
				const $tab = $(`<div class="tab">${item.label}</div>`);
				$tabs.append($tab);
				$tab.on('click', () => {
					if ($tab.hasClass('active')) return;
					$('.tab').removeClass('active');
					$tab.addClass('active');
					activeTab = item;
					$('#tabContent').empty();
					getProductList(item);
				});
				if (index === 0) {
					$tab.addClass('active');
					activeTab = item;
					getProductList(item);
				}
			});

			// 动态添加样式
			const style = document.createElement('style');
			style.innerHTML = `
        #tabs {
          display: flex;
          justify-content: center;
          margin-top: 20px;
        }
        .tab {
          padding: 10px 20px;
          margin: 0 10px;
          cursor: pointer;
          border-bottom: 2px solid transparent;
        }
        .tab.active {
          border-bottom: 2px solid rgb(0, 123, 255);
        }
        #tabContent {
          margin-top: 20px;
          background-color: #f9f9f9;
          border-radius: 10px;
          box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
          flex: 1;
          overflow-y: auto;
        }
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
        table {
          width: 100%;
          border-collapse: collapse;
        }
        th, td {
          padding: 10px;
          border: 1px solid #ddd;
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
          top: -1px;
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
      `;
			document.head.appendChild(style);

			// 关闭弹出框
			$('#popupClose').on('click', () => {
				$popup.remove();
			});
		});
	};

	const getList = () => {
		const scriptArr = document.querySelectorAll('script');
		for (let item of scriptArr) {
			if (item.dataset.suspenseModuleUuid === '6622646540') {
				let dataStr = item.innerHTML.split(' = ')[1].split('window.moduleHtml_6622646540')[0];
				// 去除多余的字符，确保JSON格式正确
				dataStr = dataStr.trim().replace(/;$/, '');
				try {
					const dataObj = JSON.parse(dataStr);
					const { _serverData, params } = dataObj._fdl_request.requestList[0];
					list = _serverData.list;
					modelId = JSON.parse(params).modelId;
					createDom();
				} catch (error) {
					console.error('Error parsing JSON:', error);
				}
			}
		}
	};

	getList();
});

// https://insights.alibaba.com/openservice/gatewayService?callback=jsonp_1741054216615_14614&modelId=10367&cardId=201268094&cardType=101002747
