import os
import webbrowser
import json

class GraphGenerator:
    def __init__(self, nodes_data, edges_data):
        self.nodes_data = nodes_data
        self.edges_data = edges_data
        self.output_filename = "dependency_graph.html"

    def generate_graph(self):
        """
        生成完全獨立的 HTML 文件，所有資源都內嵌，可在受限瀏覽器中使用
        """
        self._calculate_node_positions()
        html_content = self._generate_standalone_html()
        final_file_path = os.path.join(os.getcwd(), self.output_filename)
        
        try:
            with open(final_file_path, 'w', encoding='utf-8', errors='replace') as f:
                f.write(html_content)
            print(f"Successfully generated standalone graph at: {final_file_path}")
        except Exception as e:
            print(f"Error saving file: {e}")
            return
        webbrowser.open(f"file://{final_file_path}")

    def _generate_standalone_html(self):
        """
        生成完全獨立的 HTML，包含所有內嵌資源
        """
        processed_nodes = []
        for node in self.nodes_data:
            full_addr = node.get("full_address_label", "")
            shortest_addr = full_addr
            last_bracket_index = full_addr.rfind(']')
            if last_bracket_index != -1:
                shortest_addr = full_addr[last_bracket_index + 1:]
                if shortest_addr.startswith("'"):
                    shortest_addr = shortest_addr[1:]
            
            # === 修改：使用從 converter 傳來的 has_resolved 旗標 ===
            has_dynamic_function = node.get("has_resolved", False)
            
            # === 處理resolved_formula - 只處理路徑部分，保留公式內容 ===
            resolved_formula = node.get("resolved_formula", "")
            short_resolved_formula = resolved_formula
            full_resolved_formula = resolved_formula
            
            # 如果resolved包含路徑，創建簡短版本（只影響路徑顯示）
            if resolved_formula and ('[' in resolved_formula or '!' in resolved_formula):
                # 簡短版本：移除檔案路徑，但保留工作表引用
                if ']' in resolved_formula and '[' in resolved_formula:
                    # 處理 [file]sheet!cell 格式 -> sheet!cell
                    last_bracket = resolved_formula.rfind(']')
                    if last_bracket != -1:
                        temp_formula = resolved_formula[last_bracket + 1:]
                        if temp_formula.startswith("'"):
                            temp_formula = temp_formula[1:]
                        short_resolved_formula = temp_formula
                    else:
                        short_resolved_formula = resolved_formula
                else:
                    # 如果沒有檔案引用，保持原樣
                    short_resolved_formula = resolved_formula
            
            processed_nodes.append({
                "id": self._safe_string(node["id"]),
                "label": self._safe_string(node["label"]),
                "title": self._safe_string(node["title"]),
                "color": node["color"],
                "shape": "box",
                "x": node.get('x', 0),
                "y": node.get('y', 0),
                "level": node.get('level', 0),
                "fixed": False,
                "font": {"color": "black"},
                "filename": self._safe_string(node.get('filename', 'Current File')),
                "short_address_label": self._safe_string(node["short_address_label"]),
                "full_address_label": self._safe_string(node["full_address_label"]),
                "shortest_address_label": self._safe_string(shortest_addr),
                "short_formula_label": self._safe_string(node["short_formula_label"]),
                "full_formula_label": self._safe_string(node["full_formula_label"]),
                "value_label": self._safe_string(node["value_label"]),
                "resolved_formula": self._safe_string(resolved_formula),
                "short_resolved_formula": self._safe_string(short_resolved_formula),
                "full_resolved_formula": self._safe_string(full_resolved_formula),
                "has_dynamic_function": has_dynamic_function
            })
        
        processed_edges = []
        for edge in self.edges_data:
            processed_edges.append({
                "arrows": "to",
                "from": self._safe_string(edge[0]),
                "to": self._safe_string(edge[1])
            })
            
        nodes_json = self._safe_json_encode(processed_nodes)
        edges_json = self._safe_json_encode(processed_edges)

        print(f"Processing {len(processed_nodes)} nodes and {len(processed_edges)} edges")

        vis_js_content = """
        // Complete vis.js implementation for network visualization
        var vis = (function() {
            
            function DataSet(data) {
                this.data = data || [];
                this.length = this.data.length;
                this.initialData = JSON.parse(JSON.stringify(data));
            }
            
            DataSet.prototype.get = function(options) {
                if (options && options.returnType === "Object") {
                    var result = {};
                    this.data.forEach(item => {
                        result[item.id] = item;
                    });
                    return result;
                } else if (options && options.returnType === "Array") {
                    return this.data.slice();
                }
                return this.data.slice();
            };

            DataSet.prototype.getInitialData = function() {
                return JSON.parse(JSON.stringify(this.initialData));
            };
            
            DataSet.prototype.update = function(updates) {
                if (!Array.isArray(updates)) {
                    updates = [updates];
                }
                
                updates.forEach(update => {
                    var index = this.data.findIndex(item => item.id === update.id);
                    if (index !== -1) {
                        Object.assign(this.data[index], update);
                    }
                });
            };
            
            function Network(container, data, options) {
                this.container = container;
                this.nodes = data.nodes;
                this.edges = data.edges;
                this.options = options || {};
                this.canvas = null;
                this.ctx = null;
                this.nodePositions = {};
                this.nodeSizes = {};
                this.isDragging = false;
                this.dragNode = null;
                this.dragOffset = {x: 0, y: 0};
                this.viewOffset = {x: 0, y: 0};
                this.isDraggingView = false;
                this.lastMousePos = {x: 0, y: 0};
                this.scale = 1;
                
                this.init();
            }
            
            Network.prototype.init = function() {
                this.canvas = document.createElement('canvas');
                this.canvas.width = this.container.clientWidth;
                this.canvas.height = this.container.clientHeight;
                this.canvas.style.display = 'block';
                this.canvas.style.cursor = 'grab';
                this.container.appendChild(this.canvas);
                this.ctx = this.canvas.getContext('2d');
                
                window.addEventListener('resize', () => {
                    this.canvas.width = this.container.clientWidth;
                    this.canvas.height = this.container.clientHeight;
                    this.reorganizeLayout();
                });
                
                var nodes = this.nodes.get();
                nodes.forEach(node => {
                    this.nodePositions[node.id] = {
                        x: (node.x || 0),
                        y: (node.y || 0)
                    };
                    this.calculateNodeSize(node);
                });
                
                this.reorganizeLayout();
                this.setupEvents();
            };
            
            Network.prototype.calculateNodeSize = function(node) {
                var tempCtx = this.ctx;
                var fontSize = (node.font && node.font.size) || 14;
                var label = node.label || node.id;
                var lines = label.split('\\n');
                
                var padding = 10;
                var minHeight = 40;
                var maxContentWidth = 450;

                var textBlockHeight = 0;
                var actualContentWidth = 0;

                lines.forEach((line, index) => {
                    var cleanLine = line.replace(/<[^>]*>/g, '');
                    var fontStyle = (line.includes('<b>') ? 'bold ' : '') + (line.includes('<i>') ? 'italic ' : '');
                    tempCtx.font = fontStyle + fontSize + 'px Arial';
                    
                    if (cleanLine.trim() === '') {
                        textBlockHeight += (fontSize + 4) / 2;
                        return;
                    }

                    var words = cleanLine.split(' ');
                    var currentLine = '';
                    for (var i = 0; i < words.length; i++) {
                        var testLine = currentLine + words[i] + ' ';
                        var metrics = tempCtx.measureText(testLine);
                        
                        if (metrics.width > maxContentWidth && i > 0) {
                            actualContentWidth = Math.max(actualContentWidth, tempCtx.measureText(currentLine).width);
                            textBlockHeight += (fontSize + 4);
                            currentLine = words[i] + ' ';
                        } else {
                            currentLine = testLine;
                        }
                    }
                    
                    actualContentWidth = Math.max(actualContentWidth, tempCtx.measureText(currentLine).width);
                    textBlockHeight += (fontSize + 4);

                    if (index < lines.length - 1 && cleanLine.trim() !== '') {
                        textBlockHeight += 4;
                    }
                });
                
                textBlockHeight -= 4;

                this.nodeSizes[node.id] = {
                    width: actualContentWidth + padding * 2,
                    height: Math.max(minHeight, textBlockHeight + padding * 2),
                    textBlockHeight: textBlockHeight
                };
            };
            
            Network.prototype.reorganizeLayout = function() {
                var verticalSpacingSlider = document.getElementById('verticalSpacingSlider');
                var level_y_step = verticalSpacingSlider ? parseInt(verticalSpacingSlider.value) : 250;

                var initialNodes = this.nodes.getInitialData();
                var levels = {};
                var horizontalGap = 40;

                initialNodes.forEach(node => {
                    var level = node.level;
                    if (!levels[level]) { levels[level] = []; }
                    levels[level].push(node.id);
                });

                for (var level in levels) {
                    var levelNodes = levels[level];
                    levelNodes.sort((a, b) => {
                        var nodeA_initial = initialNodes.find(n => n.id === a);
                        var nodeB_initial = initialNodes.find(n => n.id === b);
                        return nodeA_initial.x - nodeB_initial.x;
                    });
                    
                    if (levelNodes.length > 0) {
                        var totalLevelWidth = 0;
                        levelNodes.forEach((nodeId, index) => {
                            totalLevelWidth += this.nodeSizes[nodeId].width;
                            if (index > 0) {
                                totalLevelWidth += horizontalGap;
                            }
                        });
                        
                        var startX = (this.canvas.width - totalLevelWidth) / 2;
                        var rightmostX = startX - horizontalGap;

                        for (var i = 0; i < levelNodes.length; i++) {
                            var currNodeId = levelNodes[i];
                            var currPos = this.nodePositions[currNodeId];
                            var currSize = this.nodeSizes[currNodeId];
                            
                            currPos.y = level * level_y_step;
                            
                            var requiredX = rightmostX + (currSize.width / 2) + horizontalGap;
                            currPos.x = requiredX;
                            
                            rightmostX = currPos.x + (currSize.width / 2);
                        }
                    }
                }
                this.draw();
            };
            
            Network.prototype.draw = function() {
                var ctx = this.ctx;
                ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
                ctx.save();
                ctx.scale(this.scale, this.scale);
                ctx.translate(this.viewOffset.x, this.viewOffset.y);
                
                var edges = this.edges.get();
                edges.forEach(edge => {
                    var fromPos = this.nodePositions[edge.from];
                    var toPos = this.nodePositions[edge.to];
                    if (fromPos && toPos) {
                        var edgeId = edge.from + '-' + edge.to;
                        var isHighlighted = this.highlightedEdges && this.highlightedEdges.has(edgeId);
                        
                        if (isHighlighted) {
                            ctx.shadowColor = '#FFD700';
                            ctx.shadowBlur = 8;
                            ctx.shadowOffsetX = 0;
                            ctx.shadowOffsetY = 0;
                            ctx.strokeStyle = '#000000';
                            ctx.lineWidth = 4;
                        } else {
                            ctx.shadowColor = 'transparent';
                            ctx.shadowBlur = 0;
                            ctx.strokeStyle = '#848484';
                            ctx.lineWidth = 1;
                        }
                        
                        ctx.beginPath();
                        ctx.moveTo(fromPos.x, fromPos.y);
                        ctx.lineTo(toPos.x, toPos.y);
                        ctx.stroke();
                        
                        ctx.shadowColor = 'transparent';
                        ctx.shadowBlur = 0;
                        
                        this.drawArrow(ctx, fromPos.x, fromPos.y, toPos.x, toPos.y, edge.from, edge.to, isHighlighted);
                    }
                });
                
                var nodes = this.nodes.get();
                nodes.forEach(node => {
                    var pos = this.nodePositions[node.id];
                    if (pos) {
                        var isHighlighted = this.highlightedNodes && this.highlightedNodes.has(node.id);
                        this.drawNode(ctx, node, pos.x, pos.y, isHighlighted);
                    }
                });
                ctx.restore();
            };
            
            Network.prototype.drawNode = function(ctx, node, x, y, isHighlighted) {
                var nodeSize = this.nodeSizes[node.id];
                if (!nodeSize) {
                    this.calculateNodeSize(node);
                    nodeSize = this.nodeSizes[node.id];
                }
                
                var width = nodeSize.width;
                var height = nodeSize.height;
                
                if (isHighlighted) {
                    ctx.shadowColor = '#FFD700';
                    ctx.shadowBlur = 15;
                    ctx.shadowOffsetX = 0;
                    ctx.shadowOffsetY = 0;
                    
                    var originalColor = node.color || '#97C2FC';
                    var brightColor = this.brightenColor(originalColor, 0.4);
                    ctx.fillStyle = brightColor;
                } else {
                    ctx.shadowColor = 'transparent';
                    ctx.shadowBlur = 0;
                    ctx.fillStyle = node.color || '#97C2FC';
                }
                
                ctx.fillRect(x - width/2, y - height/2, width, height);
                
                ctx.shadowColor = 'transparent';
                ctx.shadowBlur = 0;
                
                if (isHighlighted) {
                    ctx.strokeStyle = '#000000';
                    ctx.lineWidth = 4;
                } else {
                    ctx.strokeStyle = '#2B7CE9';
                    ctx.lineWidth = 1;
                }
                ctx.strokeRect(x - width/2, y - height/2, width, height);
                
                ctx.fillStyle = 'black';
                var fontSize = (node.font && node.font.size) || 14;
                ctx.textAlign = 'left';
                ctx.textBaseline = 'top';
                
                var label = node.label || node.id;
                var lines = label.split('\\n');
                var lineHeight = fontSize + 4;
                var padding = 10;
                
                var startY = y - nodeSize.textBlockHeight / 2;
                var startX = x - width/2 + padding;
                var maxLineWidth = width - padding * 2;
                
                var currentY = startY;
                
                lines.forEach((line, index) => {
                    var cleanLine = line.replace(/<b>(.*?)<\\/b>/g, '$1').replace(/<i>(.*?)<\\/i>/g, '$1').replace(/<[^>]*>/g, '');
                    var isBold = line.includes('<b>');
                    var isItalic = line.includes('<i>');
                    
                    var fontStyle = '';
                    if (isBold && isItalic) {
                        fontStyle = 'bold italic ';
                    } else if (isBold) {
                        fontStyle = 'bold ';
                    } else if (isItalic) {
                        fontStyle = 'italic ';
                    }
                    ctx.font = fontStyle + fontSize + 'px Arial';
                    
                    if (cleanLine.trim() === '') {
                        currentY += lineHeight / 2;
                        return;
                    }

                    var words = cleanLine.split(' ');
                    var currentLine = '';
                    
                    for (var i = 0; i < words.length; i++) {
                        var testLine = currentLine + words[i] + ' ';
                        var metrics = ctx.measureText(testLine);
                        var testWidth = metrics.width;
                        
                        if (testWidth > maxLineWidth && i > 0) {
                            ctx.fillText(currentLine.trim(), startX, currentY);
                            currentLine = words[i] + ' ';
                            currentY += lineHeight;
                        } else {
                            currentLine = testLine;
                        }
                    }
                    
                    if (currentLine.trim()) {
                        ctx.fillText(currentLine.trim(), startX, currentY);
                        currentY += lineHeight;
                    }
                    
                    if (index < lines.length - 1 && cleanLine.trim() !== '') {
                        currentY += 4;
                    }
                });
            };
            
            Network.prototype.drawArrow = function(ctx, fromX, fromY, toX, toY, fromNodeId, toNodeId, isHighlighted) {
                var angle = Math.atan2(toY - fromY, toX - fromX);
                var length = 10;
                
                var toNodeSize = this.nodeSizes[toNodeId];
                if (!toNodeSize) return;

                var nodeWidth = toNodeSize.width;
                var nodeHeight = toNodeSize.height;
                
                var dx = toX - fromX;
                var dy = toY - fromY;
                
                var t = 1;
                if (Math.abs(dx) * nodeHeight > Math.abs(dy) * nodeWidth) {
                    t = Math.abs(nodeWidth / (2 * dx));
                } else {
                    t = Math.abs(nodeHeight / (2 * dy));
                }
                
                var arrowX = fromX + dx * (1 - t);
                var arrowY = fromY + dy * (1 - t);

                ctx.beginPath();
                ctx.moveTo(arrowX, arrowY);
                ctx.lineTo(arrowX - length * Math.cos(angle - Math.PI / 6), 
                          arrowY - length * Math.sin(angle - Math.PI / 6));
                ctx.moveTo(arrowX, arrowY);
                ctx.lineTo(arrowX - length * Math.cos(angle + Math.PI / 6), 
                          arrowY - length * Math.sin(angle + Math.PI / 6));
                
                if (isHighlighted) {
                    ctx.strokeStyle = '#000000';
                    ctx.lineWidth = 4;
                } else {
                    ctx.strokeStyle = '#848484';
                    ctx.lineWidth = 1;
                }
                ctx.stroke();
            };
            
            Network.prototype.setupEvents = function() {
                var self = this;
                
                this.currentHighlightedNode = null;
                this.originalNodeStyles = new Map();
                this.originalEdgeStyles = new Map();
                this.isClickForHighlight = false;
                
                this.saveOriginalStyles();
                
                this.canvas.addEventListener('mousedown', function(e) {
                    var rect = self.canvas.getBoundingClientRect();
                    var mouseX = (e.clientX - rect.left) / self.scale - self.viewOffset.x;
                    var mouseY = (e.clientY - rect.top) / self.scale - self.viewOffset.y;
                    
                    self.lastMousePos = {x: e.clientX - rect.left, y: e.clientY - rect.top};
                    
                    var nodes = self.nodes.get();
                    var nodeClicked = false;
                    var clickedNodeId = null;
                    
                    for (var i = 0; i < nodes.length; i++) {
                        var node = nodes[i];
                        var pos = self.nodePositions[node.id];
                        var nodeSize = self.nodeSizes[node.id];
                        
                        if (!nodeSize) continue;
                        
                        if (pos && 
                            mouseX >= pos.x - nodeSize.width/2 && mouseX <= pos.x + nodeSize.width/2 &&
                            mouseY >= pos.y - nodeSize.height/2 && mouseY <= pos.y + nodeSize.height/2) {
                            
                            clickedNodeId = node.id;
                            nodeClicked = true;
                            
                            self.dragNode = node.id;
                            self.dragOffset = {
                                x: mouseX - pos.x,
                                y: mouseY - pos.y
                            };
                            
                            if (e.button === 0 && !e.ctrlKey && !e.shiftKey && !e.altKey) {
                                self.isClickForHighlight = true;
                                setTimeout(function() {
                                    if (self.isClickForHighlight && !self.isDragging) {
                                        self.handleNodeHighlight(clickedNodeId);
                                    }
                                }, 150);
                            }
                            break;
                        }
                    }
                    
                    if (!nodeClicked) {
                        if (e.button === 0) {
                            self.clearHighlight();
                            self.currentHighlightedNode = null;
                        }
                        self.isDraggingView = true;
                        self.canvas.style.cursor = 'grabbing';
                    }
                });
                
                this.canvas.addEventListener('mousemove', function(e) {
                    var rect = self.canvas.getBoundingClientRect();
                    var mouseX = (e.clientX - rect.left) / self.scale - self.viewOffset.x;
                    var mouseY = (e.clientY - rect.top) / self.scale - self.viewOffset.y;
                    
                    if (self.dragNode && !self.isDragging) {
                        var moveDistance = Math.sqrt(
                            Math.pow(e.clientX - rect.left - self.lastMousePos.x, 2) + 
                            Math.pow(e.clientY - rect.top - self.lastMousePos.y, 2)
                        );
                        if (moveDistance > 3) {
                            self.isDragging = true;
                            self.isClickForHighlight = false;
                            self.canvas.style.cursor = 'grabbing';
                        }
                    }
                    
                    if (self.isDragging && self.dragNode) {
                        self.nodePositions[self.dragNode] = {
                            x: mouseX - self.dragOffset.x,
                            y: mouseY - self.dragOffset.y
                        };
                        self.draw();
                    } else if (self.isDraggingView) {
                        var currentMousePos = {x: e.clientX - rect.left, y: e.clientY - rect.top};
                        var deltaX = (currentMousePos.x - self.lastMousePos.x) / self.scale;
                        var deltaY = (currentMousePos.y - self.lastMousePos.y) / self.scale;
                        
                        self.viewOffset.x += deltaX;
                        self.viewOffset.y += deltaY;
                        
                        self.draw();
                        self.lastMousePos = currentMousePos;
                    }
                });
                
                this.canvas.addEventListener('mouseup', function(e) {
                    setTimeout(function() {
                        self.isClickForHighlight = false;
                    }, 200);
                    
                    self.isDragging = false;
                    self.isDraggingView = false;
                    self.dragNode = null;
                    self.canvas.style.cursor = 'grab';
                });
                
                this.canvas.addEventListener('wheel', function(e) {
                    e.preventDefault();
                    
                    var rect = self.canvas.getBoundingClientRect();
                    var mouseX = e.clientX - rect.left;
                    var mouseY = e.clientY - rect.top;
                    
                    var scaleFactor = e.deltaY > 0 ? 0.9 : 1.1;
                    var newScale = self.scale * scaleFactor;
                    
                    newScale = Math.max(0.1, Math.min(5, newScale));
                    
                    if (newScale !== self.scale) {
                        var mousePoint = {
                            x: (mouseX / self.scale) - self.viewOffset.x,
                            y: (mouseY / self.scale) - self.viewOffset.y
                        };
                        self.scale = newScale;
                        var newMousePoint = {
                            x: (mouseX / self.scale) - self.viewOffset.x,
                            y: (mouseY / self.scale) - self.viewOffset.y
                        };
                        self.viewOffset.x += newMousePoint.x - mousePoint.x;
                        self.viewOffset.y += newMousePoint.y - mousePoint.y;
                        self.draw();
                    }
                });
            };
            
            Network.prototype.saveOriginalStyles = function() {
                var nodes = this.nodes.get();
                var edges = this.edges.get();
                
                for (var i = 0; i < nodes.length; i++) {
                    var node = nodes[i];
                    this.originalNodeStyles.set(node.id, {
                        color: node.color || '#97C2FC',
                        borderWidth: 1,
                        borderColor: '#2B7CE9'
                    });
                }
                
                for (var i = 0; i < edges.length; i++) {
                    var edge = edges[i];
                    var edgeId = edge.from + '-' + edge.to;
                    this.originalEdgeStyles.set(edgeId, {
                        color: '#848484',
                        width: 1
                    });
                }
            };
            
            Network.prototype.handleNodeHighlight = function(nodeId) {
                if (this.currentHighlightedNode === nodeId) {
                    this.clearHighlight();
                    this.currentHighlightedNode = null;
                } else {
                    this.clearHighlight();
                    this.highlightNodeAndRelated(nodeId);
                    this.currentHighlightedNode = nodeId;
                }
            };
            
            Network.prototype.highlightNodeAndRelated = function(nodeId) {
                var allEdges = this.edges.get();
                var relatedEdges = [];
                var relatedNodeIds = new Set();
                
                relatedNodeIds.add(nodeId);
                
                for (var i = 0; i < allEdges.length; i++) {
                    var edge = allEdges[i];
                    if (edge.from === nodeId || edge.to === nodeId) {
                        relatedEdges.push(edge);
                        
                        if (edge.from === nodeId) {
                            relatedNodeIds.add(edge.to);
                        }
                        if (edge.to === nodeId) {
                            relatedNodeIds.add(edge.from);
                        }
                    }
                }
                
                this.highlightNodes(Array.from(relatedNodeIds));
                this.highlightEdges(relatedEdges);
                this.draw();
            };
            
            Network.prototype.highlightNodes = function(nodeIds) {
                this.highlightedNodes = new Set(nodeIds);
            };
            
            Network.prototype.highlightEdges = function(edgesToHighlight) {
                this.highlightedEdges = new Set();
                for (var i = 0; i < edgesToHighlight.length; i++) {
                    var edge = edgesToHighlight[i];
                    var edgeId = edge.from + '-' + edge.to;
                    this.highlightedEdges.add(edgeId);
                }
            };
            
            Network.prototype.clearHighlight = function() {
                this.highlightedNodes = new Set();
                this.highlightedEdges = new Set();
                this.draw();
            };
            
            Network.prototype.brightenColor = function(color, factor) {
                var hex = color.replace('#', '');
                var r = parseInt(hex.substr(0, 2), 16);
                var g = parseInt(hex.substr(2, 2), 16);
                var b = parseInt(hex.substr(4, 2), 16);
                
                r = Math.min(255, Math.floor(r + (255 - r) * factor));
                g = Math.min(255, Math.floor(g + (255 - g) * factor));
                b = Math.min(255, Math.floor(b + (255 - b) * factor));
                
                var rHex = r.toString(16).padStart(2, '0');
                var gHex = g.toString(16).padStart(2, '0');
                var bHex = b.toString(16).padStart(2, '0');
                
                return '#' + rHex + gHex + bHex;
            };
            
            return {
                DataSet: DataSet,
                Network: Network
            };
        })();
        """

        html_template = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Dependency Graph - Standalone</title>
    
    <style type="text/css">
        body {{
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            background-color: #f5f5f5;
        }}
        
        #mynetwork {{
            width: 100%;
            height: 100vh;
            background-color: #ffffff;
            border: 1px solid lightgray;
            position: relative;
        }}
        
        .controls {{
            position: absolute;
            top: 10px;
            left: 10px;
            background: rgba(248, 249, 250, 0.95);
            padding: 12px;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            z-index: 1000;
            font-family: sans-serif;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            width: 360px;
            max-height: 90vh;
            overflow-y: auto;
        }}
        
        .controls h4 {{
            margin: 0 0 10px 0;
            font-weight: bold;
            color: #333;
        }}
        
        .control-item {{
            margin-bottom: 8px;
        }}
        
        .control-item label {{
            cursor: pointer;
            display: flex;
            align-items: center;
        }}

        .control-item label.disabled {{
            color: #888;
            cursor: not-allowed;
        }}
        
        .control-item input[type="checkbox"] {{
            margin-right: 8px;
        }}
        
        .slider-container {{
            margin-bottom: 8px;
        }}
        
        .slider-container label {{
            display: block;
            margin-bottom: 4px;
            cursor: pointer;
        }}
        
        .slider-container input[type="range"] {{
            width: 100%;
        }}

        #reorganizeButton {{
            width: 100%;
            padding: 8px;
            margin-top: 8px;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 1em;
            font-weight: bold;
        }}

        #reorganizeButton:hover {{
            background-color: #0056b3;
        }}
        
        .legend {{
            border-top: 1px solid #ccc;
            margin-top: 12px;
            padding-top: 10px;
        }}
        
        .legend h5 {{
            margin: 0 0 6px 0;
            font-weight: bold;
            color: #333;
        }}
        
        .legend-item {{
            display: flex;
            align-items: flex-start;
            margin-bottom: 6px;
        }}
        
        .legend-color {{
            width: 1.2em;
            height: 1.2em;
            margin-right: 8px;
            border-radius: 3px;
            border: 1px solid #ddd;
            flex-shrink: 0;
            margin-top: 0.1em;
        }}
        
        .legend-text {{
            font-weight: 500;
            color: #333;
            word-break: break-all;
        }}
        
        .legend-help {{
            margin-top: 8px;
            font-size: 0.8em;
            color: #666;
        }}
    </style>
</head>

<body>
    <!-- 控制面板 -->
    <div class="controls" id="controls-panel">
        <h4>Display Options</h4>
        
        <div class="control-item">
            <label>
                <input type='checkbox' id='hideAddressFileToggle'> Hide Address File Name
            </label>
        </div>

        <div class="control-item">
            <label id="fullAddressLabel">
                <input type='checkbox' id='addressToggle'> Show Full Address Path
            </label>
        </div>

        <div class="control-item">
            <label>
                <input type='checkbox' id='formulaToggle'> Show Full Formula Path
            </label>
        </div>
        
        <div class="control-item">
            <label>
                <input type='checkbox' id='resolvedToggle'> Show Full Resolved Path
            </label>
        </div>
        
        <div class="slider-container">
            <label for='fontSizeSlider'>
                Node Font Size: <span id='fontSizeValue'>14</span>px
            </label>
            <input type='range' id='fontSizeSlider' min='10' max='24' value='14'>
        </div>
        
        <div class="slider-container">
            <label for='lineWidthSlider'>
                Line Width: <span id='lineWidthValue'>60</span> chars
            </label>
            <input type='range' id='lineWidthSlider' min='30' max='120' value='60'>
        </div>
        
        <div class="slider-container">
            <label for='verticalSpacingSlider'>
                Vertical Spacing: <span id='verticalSpacingValue'>250</span>px
            </label>
            <input type='range' id='verticalSpacingSlider' min='0' max='800' value='250'>
        </div>
        
        <div class="slider-container">
            <label for='uiFontSizeSlider'>
                UI Font Size: <span id='uiFontSizeValue'>14</span>px
            </label>
            <input type='range' id='uiFontSizeSlider' min='10' max='24' value='14'>
        </div>

        <button id="reorganizeButton">Re-organize Layout</button>

        <div class="legend">
            <h5>File Legend</h5>
            <div id='fileLegend'>
            </div>
            <div class="legend-help">
                相同顏色 = 同一檔案<br>
                不同顏色間的箭頭 = 跨檔案依賴
            </div>
        </div>
    </div>
    
    <div id="mynetwork"></div>

    <script type="text/javascript">
        {vis_js_content}
    </script>

    <script type="text/javascript">
        var nodes;
        var edges;
        var network;
        var nodeData = {nodes_json};
        var edgeData = {edges_json};

        function initGraph() {{
            console.log('Initializing graph with', nodeData.length, 'nodes and', edgeData.length, 'edges');
            var container = document.getElementById('mynetwork');
            nodes = new vis.DataSet(nodeData);
            edges = new vis.DataSet(edgeData);
            var data = {{ nodes: nodes, edges: edges }};
            var options = {{
                interaction: {{ dragNodes: true, dragView: true, zoomView: true }},
                physics: {{ enabled: false }}
            }};
            network = new vis.Network(container, data, options);
            console.log('Graph initialized successfully');
            initControls();
        }}
        
        function initControls() {{
            var controlsPanel = document.getElementById('controls-panel');
            var hideAddressFileToggle = document.getElementById('hideAddressFileToggle');
            var addressToggle = document.getElementById('addressToggle');
            var fullAddressLabel = document.getElementById('fullAddressLabel');
            var formulaToggle = document.getElementById('formulaToggle');
            var resolvedToggle = document.getElementById('resolvedToggle');
            var fontSizeSlider = document.getElementById('fontSizeSlider');
            var fontSizeValue = document.getElementById('fontSizeValue');
            var lineWidthSlider = document.getElementById('lineWidthSlider');
            var lineWidthValue = document.getElementById('lineWidthValue');
            var verticalSpacingSlider = document.getElementById('verticalSpacingSlider');
            var verticalSpacingValue = document.getElementById('verticalSpacingValue');
            var uiFontSizeSlider = document.getElementById('uiFontSizeSlider');
            var uiFontSizeValue = document.getElementById('uiFontSizeValue');
            var reorganizeButton = document.getElementById('reorganizeButton');
            
            // === 簡化的換行函數 - 只在指定字符數處截斷 ===
            function wrapText(text, maxWidth) {{
                if (!text || text.length <= maxWidth) {{
                    return text;
                }}
                
                var result = [];
                var i = 0;
                
                while (i < text.length) {{
                    var chunk = text.substr(i, maxWidth);
                    result.push(chunk);
                    i += maxWidth;
                }}
                
                return result.join('\\n');
            }}
            
            function updateNodeLabels() {{
                if (!network || !nodes) return;
                
                console.log('Updating node labels...');
                
                var hideAddressFile = hideAddressFileToggle.checked;
                var showFullAddress = addressToggle.checked;
                var showFullFormula = formulaToggle.checked;
                var showFullResolved = resolvedToggle.checked;
                var fontSize = parseInt(fontSizeSlider.value);
                var lineWidth = parseInt(lineWidthSlider.value);
                
                var allNodes = nodes.get();
                var updatedNodes = [];
                
                allNodes.forEach(function(node) {{
                    var addressLabel;
                    if (hideAddressFile) {{
                        addressLabel = node.shortest_address_label || node.short_address_label;
                    }} else {{
                        addressLabel = showFullAddress ? node.full_address_label : node.short_address_label;
                    }}

                    var formulaLabel = showFullFormula ? node.full_formula_label : node.short_formula_label;
                    
                    var newLabel = 'Address : <b>' + (addressLabel || node.short_address_label) + '</b>';
                    
                    // === 套用簡化換行到Formula ===
                    if (formulaLabel && formulaLabel !== 'N/A' && formulaLabel !== null) {{
                        var displayFormula = formulaLabel.startsWith('=') ? formulaLabel : '=' + formulaLabel;
                        displayFormula = wrapText(displayFormula, lineWidth);
                        newLabel += '\\n\\nFormula : <i>' + displayFormula + '</i>';
                    }} else {{
                        newLabel += '\\n\\nFormula : <i>N/A</i>';
                    }}
                    
                    // === 套用簡化換行到Resolved ===
                    var shouldShowResolved = false;
                    var resolvedToUse = '';
                    
                    if (node.resolved_formula && node.resolved_formula !== 'N/A' && node.resolved_formula !== null && node.resolved_formula !== formulaLabel) {{
                        if (node.has_dynamic_function) {{
                            shouldShowResolved = true;
                            resolvedToUse = showFullResolved ? node.full_resolved_formula : node.short_resolved_formula;
                        }} else if (showFullResolved) {{
                            shouldShowResolved = true;
                            resolvedToUse = node.full_resolved_formula;
                        }}
                    }}
                    
                    if (shouldShowResolved && resolvedToUse) {{
                        var displayResolved = resolvedToUse.startsWith('=') ? resolvedToUse : '=' + resolvedToUse;
                        displayResolved = wrapText(displayResolved, lineWidth);
                        newLabel += '\\n\\nResolved : <i>' + displayResolved + '</i>';
                    }}
                    
                    newLabel += '\\n\\nValue     : ' + (node.value_label || 'N/A');
                    
                    updatedNodes.push({{
                        id: node.id,
                        label: newLabel,
                        font: {{ size: fontSize }}
                    }});
                }});
                
                if (updatedNodes.length > 0) {{
                    nodes.update(updatedNodes);
                    allNodes.forEach(node => network.calculateNodeSize(node));
                    network.reorganizeLayout();
                }}
                
                console.log('Node labels updated with simple line wrapping, line width:', lineWidth);
            }}
            
            function updateNodeFontSize() {{
                fontSizeValue.textContent = fontSizeSlider.value;
                updateNodeLabels();
            }}

            function updateLineWidth() {{
                lineWidthValue.textContent = lineWidthSlider.value;
                updateNodeLabels();
            }}

            function updateVerticalSpacing() {{
                verticalSpacingValue.textContent = verticalSpacingSlider.value;
                network.reorganizeLayout();
            }}

            function updateUiFontSize() {{
                var newSize = uiFontSizeSlider.value + 'px';
                uiFontSizeValue.textContent = uiFontSizeSlider.value;
                controlsPanel.style.fontSize = newSize;
            }}

            function handleReorganize() {{
                if (network) {{
                    console.log("Re-organizing layout...");
                    network.reorganizeLayout();
                }}
            }}
            
            function generateFileLegend() {{
                var fileLegendDiv = document.getElementById('fileLegend');
                if (!fileLegendDiv || !nodes) return;
                var fileColors = new Map();
                var allNodes = nodes.get();
                allNodes.forEach(function(node) {{
                    var color = node.color || '#808080';
                    var filename = node.filename || 'Unknown File';
                    if (!fileColors.has(filename)) {{
                        fileColors.set(filename, color);
                    }}
                }});
                var sortedFiles = Array.from(fileColors.entries()).sort(function(a, b) {{
                    if (a[0] === 'Current File') return -1;
                    if (b[0] === 'Current File') return 1;
                    return a[0].localeCompare(b[0]);
                }});
                var legendHTML = '';
                sortedFiles.forEach(function(item) {{
                    var filename = item[0];
                    var color = item[1];
                    legendHTML += '<div class="legend-item" title="檔案: ' + filename + '">';
                    legendHTML += '<div class="legend-color" style="background-color: ' + color + ';"></div>';
                    legendHTML += '<span class="legend-text">' + filename + '</span>';
                    legendHTML += '</div>';
                }});
                fileLegendDiv.innerHTML = legendHTML;
            }}
            
            // === 事件監聽器 ===
            hideAddressFileToggle.addEventListener('change', function() {{
                if (this.checked) {{
                    addressToggle.disabled = true;
                    fullAddressLabel.classList.add('disabled');
                }} else {{
                    addressToggle.disabled = false;
                    fullAddressLabel.classList.remove('disabled');
                }}
                updateNodeLabels();
            }});

            addressToggle.addEventListener('change', updateNodeLabels);
            formulaToggle.addEventListener('change', updateNodeLabels);
            resolvedToggle.addEventListener('change', updateNodeLabels);
            fontSizeSlider.addEventListener('input', updateNodeFontSize);
            lineWidthSlider.addEventListener('input', updateLineWidth);
            verticalSpacingSlider.addEventListener('input', updateVerticalSpacing);
            uiFontSizeSlider.addEventListener('input', updateUiFontSize);
            reorganizeButton.addEventListener('click', handleReorganize);
            
            generateFileLegend();
            updateUiFontSize();
            updateNodeLabels();
            
            console.log('Controls initialized with simple line wrapping');
        }}
        
        window.addEventListener('load', function() {{
            initGraph();
        }});
    </script>
</body>
</html>"""
        
        return html_template

    def _safe_string(self, value):
        if value is None:
            return ""
        
        str_value = str(value)
        try:
            return str_value.encode('utf-8', errors='ignore').decode('utf-8')
        except:
            return str_value.encode('ascii', errors='ignore').decode('ascii')

    def _safe_json_encode(self, data):
        try:
            return json.dumps(data, ensure_ascii=False, separators=(',', ':'))
        except Exception as e:
            print(f"JSON encoding error: {e}")
            return json.dumps(data, ensure_ascii=True, separators=(',', ':'))

    def _calculate_node_positions(self):
        """
        只計算節點的初始 X 座標，Y 座標由 JS 根據層級和用戶設置的間距動態計算
        """
        level_counts = {}
        for node in self.nodes_data:
            level = node.get('level', 0)
            if level not in level_counts:
                level_counts[level] = 0
            level_counts[level] += 1

        level_x_step = 450

        current_level_counts = {level: 0 for level in level_counts}

        for node in self.nodes_data:
            level = node.get('level', 0)
            total_in_level = level_counts.get(level, 1)
            current_index_in_level = current_level_counts.get(level, 0)
            
            x = (current_index_in_level - (total_in_level - 1) / 2.0) * level_x_step
            
            node['x'] = x
            node['y'] = 0
            current_level_counts[level] = current_level_counts.get(level, 0) + 1