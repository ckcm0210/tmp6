// Enhanced utils.js for fixing graph highlighting issues
// Fix graph highlighting function failure problem

var allNodes;
var allEdges;
var highlightActive = false;
var selectedNode = null;
var clickTimeout = null;

function neighbourhoodHighlight(params) {
    // Clear any existing timeout to prevent multiple rapid clicks
    if (clickTimeout) {
        clearTimeout(clickTimeout);
    }
    
    clickTimeout = setTimeout(function() {
        try {
            // Get the data from the vis DataSets
            allNodes = nodes.get({returnType:"Object"});
            allEdges = edges.get({returnType:"Object"});
            
            // If something is selected:
            if (params.nodes.length > 0) {
                var selectedNodeId = params.nodes[0];
                
                // If clicking the same node again, clear highlighting
                if (selectedNode === selectedNodeId && highlightActive) {
                    clearHighlight();
                    return;
                }
                
                selectedNode = selectedNodeId;
                highlightActive = true;
                
                var updateArray = [];
                var edgeUpdateArray = [];
                
                // Mark all nodes as hidden initially
                for (var nodeId in allNodes) {
                    allNodes[nodeId].color = 'rgba(200,200,200,0.5)';
                    allNodes[nodeId].font = {color: 'rgba(120,120,120,0.8)', size: 14};
                    allNodes[nodeId].borderWidth = 1;
                    if (allNodes[nodeId].hiddenLabel === undefined) {
                        allNodes[nodeId].hiddenLabel = allNodes[nodeId].label;
                    }
                }
                
                // Mark all edges as hidden initially
                for (var edgeId in allEdges) {
                    allEdges[edgeId].color = 'rgba(200,200,200,0.3)';
                    allEdges[edgeId].width = 1;
                }
                
                var connectedNodes = network.getConnectedNodes(selectedNodeId);
                var connectedEdges = network.getConnectedEdges(selectedNodeId);
                
                // Highlight the selected node with thick blue border
                allNodes[selectedNodeId].color = {
                    border: '#0000FF',
                    background: allNodes[selectedNodeId].originalColor || '#97C2FC',
                    highlight: {
                        border: '#0000FF',
                        background: '#CCDDFF'
                    }
                };
                allNodes[selectedNodeId].font = {color: '#000000', size: 16, face: 'arial'};
                allNodes[selectedNodeId].borderWidth = 4; // Thick border for selected node
                
                // Highlight connected nodes
                for (var i = 0; i < connectedNodes.length; i++) {
                    var nodeId = connectedNodes[i];
                    
                    // Determine if this is a precedent or dependent
                    var isPrecedent = false;
                    var isDependent = false;
                    
                    for (var j = 0; j < connectedEdges.length; j++) {
                        var edge = allEdges[connectedEdges[j]];
                        if (edge.to === selectedNodeId && edge.from === nodeId) {
                            isPrecedent = true;
                        }
                        if (edge.from === selectedNodeId && edge.to === nodeId) {
                            isDependent = true;
                        }
                    }
                    
                    if (isPrecedent) {
                        // Green for precedents (nodes that this node depends on)
                        allNodes[nodeId].color = {
                            border: '#008000',
                            background: allNodes[nodeId].originalColor || '#97C2FC',
                            highlight: {
                                border: '#008000',
                                background: '#CCFFCC'
                            }
                        };
                        allNodes[nodeId].borderWidth = 3; // Thick border
                    } else if (isDependent) {
                        // Red for dependents (nodes that depend on this node)
                        allNodes[nodeId].color = {
                            border: '#800000',
                            background: allNodes[nodeId].originalColor || '#97C2FC',
                            highlight: {
                                border: '#800000',
                                background: '#FFCCCC'
                            }
                        };
                        allNodes[nodeId].borderWidth = 3; // Thick border
                    }
                    
                    allNodes[nodeId].font = {color: '#000000', size: 14, face: 'arial'};
                }
                
                // Highlight connected edges with thick lines
                for (var i = 0; i < connectedEdges.length; i++) {
                    var edgeId = connectedEdges[i];
                    var edge = allEdges[edgeId];
                    
                    if (edge.to === selectedNodeId) {
                        // Incoming edge (precedent) - green
                        allEdges[edgeId].color = {color: '#008000', highlight: '#008000'};
                        allEdges[edgeId].width = 4; // Thick line
                    } else if (edge.from === selectedNodeId) {
                        // Outgoing edge (dependent) - red
                        allEdges[edgeId].color = {color: '#800000', highlight: '#800000'};
                        allEdges[edgeId].width = 4; // Thick line
                    }
                }
                
                // Transform the object into an array for updating
                for (var nodeId in allNodes) {
                    updateArray.push(allNodes[nodeId]);
                }
                for (var edgeId in allEdges) {
                    edgeUpdateArray.push(allEdges[edgeId]);
                }
                
                // Update the network
                nodes.update(updateArray);
                edges.update(edgeUpdateArray);
                
            } else if (highlightActive === true) {
                // Clicked on empty space, clear highlighting
                clearHighlight();
            }
        } catch (error) {
            console.error('Error in neighbourhoodHighlight:', error);
            // Reset highlighting state on error
            highlightActive = false;
            selectedNode = null;
        }
    }, 100); // Small delay to prevent rapid clicking issues
}

function clearHighlight() {
    try {
        var updateArray = [];
        var edgeUpdateArray = [];
        
        // Reset all nodes to original state
        for (var nodeId in allNodes) {
            allNodes[nodeId].color = allNodes[nodeId].originalColor || {
                border: '#2B7CE9',
                background: '#97C2FC',
                highlight: {
                    border: '#2B7CE9',
                    background: '#D2E5FF'
                }
            };
            allNodes[nodeId].font = {color: '#343434', size: 14, face: 'arial'};
            allNodes[nodeId].borderWidth = 1; // Normal border width
            if (allNodes[nodeId].hiddenLabel !== undefined) {
                allNodes[nodeId].label = allNodes[nodeId].hiddenLabel;
                allNodes[nodeId].hiddenLabel = undefined;
            }
            updateArray.push(allNodes[nodeId]);
        }
        
        // Reset all edges to original state
        for (var edgeId in allEdges) {
            allEdges[edgeId].color = allEdges[edgeId].originalColor || {
                color: '#848484',
                highlight: '#848484'
            };
            allEdges[edgeId].width = allEdges[edgeId].originalWidth || 1; // Normal line width
            edgeUpdateArray.push(allEdges[edgeId]);
        }
        
        // Update the network
        nodes.update(updateArray);
        edges.update(edgeUpdateArray);
        
        highlightActive = false;
        selectedNode = null;
        
    } catch (error) {
        console.error('Error in clearHighlight:', error);
        // Force reset state
        highlightActive = false;
        selectedNode = null;
    }
}