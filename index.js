// 你可以在这里添加更多的代码来操作宿主网页的window对象

$(document).ready(() => {
	let list = [];
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
				padding: '10%',
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

			// 请求对应的数据
			const getProductList = (item) => {
				$.ajax({
					url: `https://insights.alibaba.com/openservice/gatewayService?modelId=${modelId}&cardId=${item.cardId}&cardType=${item.cardType}&language=zh`,
					method: 'GET',
					dataType: 'json',
					success: ({ data }) => {
						const productList = data.list[0].productList;
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
			$('#exportButton').on('click', () => {
				const wb = XLSX.utils.table_to_book($('#tabContent table')[0]);
				const title = $('#popupTitle').text();
				const label = activeTab.label;
				XLSX.writeFile(wb, `${title}-${label}-数据列表.xlsx`);
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
