// 慢性アニメーションシステム
let isAnimating = false;
let animationId = null;
const ANIMATION_DAMPING = 0.15; // 減衣係数（小さいほどスムーズ）
const ANIMATION_THRESHOLD = 0.5; // 停止闾値

// ノードに目標位置プロパティを追加
function initializeNodeAnimation(node) {
    if (!node.hasOwnProperty('targetX')) {
        node.targetX = node.x || 0;
        node.targetY = node.y || 0;
        node.velocityX = 0;
        node.velocityY = 0;
    }
    if (node.children) {
        node.children.forEach(child => initializeNodeAnimation(child));
    }
}

// イージング関数（スムーズな動き）
function easeOutCubic(t) {
    return 1 - Math.pow(1 - t, 3);
}

// UUID生成関数（crypto.randomUUIDの代替）
function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        const r = Math.random() * 16 | 0;
        const v = c == 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

// ノードを目標位置に設定（慢性アニメーション付き）
function setNodeTarget(node, targetX, targetY) {
    initializeNodeAnimation(node);
    node.targetX = targetX;
    node.targetY = targetY;
}

// ノードを即座位置に移動（アニメーションなし）
function setNodeImmediately(node, x, y) {
    initializeNodeAnimation(node);
    node.x = x;
    node.y = y;
    node.targetX = x;
    node.targetY = y;
    node.velocityX = 0;
    node.velocityY = 0;
}

// アニメーションループ
function animateNodes() {
    let hasMovement = false;
    
    function animateNode(node) {
        initializeNodeAnimation(node);
        
        // 目標位置への距離を計算
        const deltaX = node.targetX - node.x;
        const deltaY = node.targetY - node.y;
        const distance = Math.sqrt(deltaX * deltaX + deltaY * deltaY);
        
        if (distance > ANIMATION_THRESHOLD) {
            // 慣性アニメーションで移動
            node.x += deltaX * ANIMATION_DAMPING;
            node.y += deltaY * ANIMATION_DAMPING;
            hasMovement = true;
        } else {
            // 目標位置に到達
            node.x = node.targetX;
            node.y = node.targetY;
        }
        
        // 子ノードを再帰的に処理
        if (node.children) {
            node.children.forEach(child => animateNode(child));
        }
    }
    
    animateNode(root);
    
    // 描画更新
    draw();
    
    // まだ動いているノードがある場合は継続
    if (hasMovement && isAnimating) {
        animationId = requestAnimationFrame(animateNodes);
    } else {
        // アニメーション終了時に衝突解消を実行
        if (isAnimating) {
            console.log('🎨 アニメーション終了 - 衝突解消開始');
            isAnimating = false;
            animationId = null;
            
            // 衝突解消を慣性アニメーションで実行
            resolveAllCollisionsWithAnimation();
        }
    }
}

// アニメーション開始
function startAnimation() {
    if (!isAnimating) {
        isAnimating = true;
        console.log('🎨 アニメーション開始');
        animateNodes();
    }
}

// アニメーション停止
function stopAnimation() {
    if (animationId) {
        cancelAnimationFrame(animationId);
        animationId = null;
    }
    isAnimating = false;
}

const svg = document.getElementById("mindmap");
if (!svg)
    throw new Error("SVG element not found");
let currentLayout = 'radial';
let mindMaps = [];
let activeMindMapId = '';

// URLパラメータからマインドマップIDを取得
function getURLParameter(name) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(name);
}

// URLで指定されたマインドマップを読み込み
function loadSpecifiedMindMap() {
    const requestedId = getURLParameter('id');
    if (requestedId) {
        console.log('URLで指定されたマインドマップID:', requestedId);
        // ローカルストレージからデータを読み込み
        try {
            const stored = localStorage.getItem('mindMaps');
            if (stored) {
                const data = JSON.parse(stored);
                mindMaps = data.mindMaps || [];
                const targetMindMap = mindMaps.find(m => m.id === requestedId);
                if (targetMindMap) {
                    activeMindMapId = targetMindMap.id;
                    root = targetMindMap.rootNode;
                    currentLayout = targetMindMap.layout || 'radial';
                    console.log('指定されたマインドマップを読み込みました:', targetMindMap.name);
                    return true;
                } else {
                    console.warn('指定されたマインドマップが見つかりません:', requestedId);
                }
            }
        } catch (error) {
            console.error('マインドマップ読み込みエラー:', error);
        }
    }
    return false;
}

// URLパラメータからマインドマップIDを取得
function getURLParameter(name) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(name);
}

// URLで指定されたマインドマップを読み込み
function loadSpecifiedMindMap() {
    const requestedId = getURLParameter('id');
    if (requestedId) {
        console.log('URLで指定されたマインドマップID:', requestedId);
        // ローカルストレージからデータを読み込み
        try {
            const stored = localStorage.getItem('mindMaps');
            if (stored) {
                const data = JSON.parse(stored);
                mindMaps = data.mindMaps || [];
                const targetMindMap = mindMaps.find(m => m.id === requestedId);
                if (targetMindMap) {
                    activeMindMapId = targetMindMap.id;
                    root = restoreTree(targetMindMap.rootNode, null); // 親情報を必ず再構築
                    currentLayout = targetMindMap.layout || 'radial';
                    console.log('指定されたマインドマップを読み込みました:', targetMindMap.name);
                    if (typeof layout === 'function') layout(root);
                    if (typeof draw === 'function') draw();
                    return true;
                } else {
                    console.warn('指定されたマインドマップが見つかりません:', requestedId);
                }
            }
        } catch (error) {
            console.error('マインドマップ読み込みエラー:', error);
        }
    }
    return false;
}
// ------------------------------
// Undo / Redo
// ------------------------------
let undoHistory = [];
let historyIndex = -1;
function pushHistory() {
    const snapshot = JSON.stringify(root, replacer);
    undoHistory = undoHistory.slice(0, historyIndex + 1);
    undoHistory.push(snapshot);
    historyIndex++;
}
function undo() {
    if (historyIndex <= 0)
        return;
    historyIndex--;
    loadSnapshot(undoHistory[historyIndex]);
    saveCurrentMindMap(); // Undo時に即座保存
}
function redo() {
    if (historyIndex >= undoHistory.length - 1)
        return;
    historyIndex++;
    loadSnapshot(undoHistory[historyIndex]);
    saveCurrentMindMap(); // Redo時に即座保存
}
function loadSnapshot(json) {
    const obj = JSON.parse(json);
    restoreTree(obj, null);
    selected = root;
    // レイアウトを即座に適用し、その後描画
    layout(root);
    draw();
}
function restoreTree(obj, parent) {
    const node = {
        id: obj.id,
        text: obj.text,
        children: [],
        parent: parent || null,
        x: obj.x,
        y: obj.y
    };
    if (parent === null) root = node;
    for (const c of obj.children) {
        const child = restoreTree(c, node);
        child.parent = node; // 明示的に親をセット
        node.children.push(child);
    }
    return node;
}
function replacer(key, value) {
    if (key === "parent")
        return undefined;
    return value;
}
// ------------------------------
// 初期データ
// ------------------------------
let root = {
    id: "root",
    text: "中心テーマ",
    children: [],
    parent: null,
    x: 800, // 新しいレイアウトの中心位置
    y: 450
};
let selected = root;
// selectedの安全な取得
function getSelected() {
    if (!selected || typeof selected !== 'object') {
        console.warn('selectedが無効です、rootに設定します');
        selected = root;
    }
    console.log('🔍 getSelected()呼び出し - 現在選択中:', selected ? selected.text : 'null', 'ID:', selected ? selected.id : 'null');
    return selected;
}
// selectedの安全な設定  
function setSelected(node) {
    console.log('🔄 setSelected()呼び出し - 新しい選択:', node ? node.text : 'null', 'ID:', node ? node.id : 'null');
    if (!node || typeof node !== 'object') {
        console.warn('無効なノードが設定されようとしました、rootに設定します');
        selected = root;
    }
    else {
        selected = node;
    }
    console.log('✅ 選択完了 - 現在選択中:', selected ? selected.text : 'null', 'ID:', selected ? selected.id : 'null');
}
// ------------------------------
// ノード操作
// ------------------------------
// 新しいノードのみを配置し、既存ノードの位置を保持
function layoutNewNodeOnly(newNode, parent) {
    const parentChildren = parent.children;
    const childIndex = parentChildren.indexOf(newNode);
    
    if (childIndex === -1) return;
    
    // 親ノードの情報を取得
    initializeNodeAnimation(parent);
    const parentX = parent.targetX || parent.x;
    const parentY = parent.targetY || parent.y;
    
    let newX, newY;
    
    // レイアウトタイプに応じて新しいノードの位置を計算
    switch (currentLayout) {
        case 'radial':
            newX = parentX + Math.cos((childIndex * 60) * Math.PI / 180) * 150;
            newY = parentY + Math.sin((childIndex * 60) * Math.PI / 180) * 150;
            break;
            
        case 'leftright':
            if (childIndex % 2 === 0) {
                newX = parentX + 180;
                newY = parentY + (childIndex - parentChildren.length / 2) * 80;
            } else {
                newX = parentX - 180;
                newY = parentY + (childIndex - parentChildren.length / 2) * 80;
            }
            break;
            
        case 'tree':
        default:
            newX = parentX + 180;
            newY = parentY + (childIndex - (parentChildren.length - 1) / 2) * 60;
            break;
    }
    
    // 新しいノードの位置を設定（親の位置からアニメーション開始）
    setNodeImmediately(newNode, parentX, parentY); // 初期位置は親と同じ
    setNodeTarget(newNode, newX, newY); // 目標位置を設定
    
    // アニメーション開始
    startAnimation();
}

// ノード削除後の最小限のレイアウト調整
function adjustLayoutAfterDeletion(parentNode) {
    if (!parentNode || !parentNode.children || parentNode.children.length === 0) {
        return;
    }
    
    // 親ノードの子ノードのみ再配置
    const children = parentNode.children;
    const parentX = parentNode.targetX || parentNode.x;
    const parentY = parentNode.targetY || parentNode.y;
    
    children.forEach((child, index) => {
        let newX, newY;
        
        switch (currentLayout) {
            case 'radial':
                const angle = (index * (360 / children.length)) * Math.PI / 180;
                newX = parentX + Math.cos(angle) * 150;
                newY = parentY + Math.sin(angle) * 150;
                break;
                
            case 'leftright':
                if (index % 2 === 0) {
                    newX = parentX + 180;
                    newY = parentY + (index - children.length / 2) * 80;
                } else {
                    newX = parentX - 180;
                    newY = parentY + (index - children.length / 2) * 80;
                }
                break;
                
            case 'tree':
            default:
                newX = parentX + 180;
                newY = parentY + (index - (children.length - 1) / 2) * 60;
                break;
        }
        
        setNodeTarget(child, newX, newY);
    });
    
    // アニメーション開始
    startAnimation();
}

function addNode(parent) {
    const node = {
        id: generateUUID(),
        text: "新しいノード",
        children: [],
        parent,
        x: parent.x, // 初期位置は親ノードと同じ
        y: parent.y
    };
    parent.children.push(node);
    pushHistory();
    
    // 新しいノードのみを配置（既存ノードの位置を保持）
    setTimeout(() => {
        layoutNewNodeOnly(node, parent);
        // レイアウト完了後にY座標順で整理
        setTimeout(() => reorganizeNodesByYPosition(), 100);
    }, 50);
    return node;
}
function deleteNode(node) {
    console.log('🗑️ 削除要求開始:', node ? node.text : 'null', 'ID:', node ? node.id : 'null', '親ノード:', node?.parent ? node.parent.text : 'null');
    
    if (!node) {
        console.warn('⚠️ 削除対象のノードが存在しません');
        alert('削除するノードが選択されていません。');
        return;
    }
    
    // 真のルートノード（idがrootのもの）かチェック
    if (node.id === 'root' || node === root) {
        console.warn('⚠️ 真のルートノードの削除は禁止されています:', node.text);
        alert(`ルートノード「${node.text}」は削除できません。\n子ノードを選択してから削除してください。`);
        return;
    }
    
    let parent = node.parent;
    
    // parentプロパティが設定されていない場合、全体から親ノードを検索
    if (!parent) {
        console.log('⚠️ parentプロパティが設定されていません。親ノードを検索します...');
        parent = findParentNode(root, node);
        if (parent) {
            console.log('✅ 親ノードを発見:', parent.text);
            // 親子関係を修復
            node.parent = parent;
        }
    }
    
    if (!parent) {
        console.warn('⚠️ 親ノードが見つかりません。このノードは削除できません:', node.text);
        alert(`ノード「${node.text}」の親ノードが見つかりません。\nデータ構造に問題がある可能性があります。`);
        return;
    }
    
    console.log('📝 削除処理実行中 - ノード:', node.text, '親:', parent.text, '親の子配列サイズ:', parent.children.length);
    
    // 削除実行
    parent.children = parent.children.filter(c => c !== node);
    
    // 次に選択するノードを決定
    if (parent.children.length > 0) {
        // 兄弟ノードがある場合は最初の兄弟を選択
        selected = parent.children[0];
    } else {
        // 兄弟ノードがない場合は親を選択
        selected = parent;
    }
    
    console.log('✅ ノード削除成功:', node.text, '親ノード:', parent.text, '次の選択:', selected.text);
    pushHistory();
    
    // 親ノードの子ノードのみ再配置（他のノードの位置を保持）
    setTimeout(() => {
        adjustLayoutAfterDeletion(parent);
        // 削除後にY座標順で整理
        setTimeout(() => reorganizeNodesByYPosition(), 100);
    }, 50);
}

// 全体のノードツリーから指定ノードの親を検索する関数
function findParentNode(searchRoot, targetNode) {
    if (!searchRoot || !targetNode) return null;
    
    // searchRootの子ノードの中にtargetNodeがあるかチェック
    if (searchRoot.children) {
        for (const child of searchRoot.children) {
            if (child.id === targetNode.id) {
                console.log('🔍 親ノード発見:', searchRoot.text, '-> 子ノード:', child.text);
                return searchRoot;
            }
        }
        
        // 再帰的に子ノード以下を検索
        for (const child of searchRoot.children) {
            const found = findParentNode(child, targetNode);
            if (found) return found;
        }
    }
    
    return null;
}

// 折りたたみ後のレイアウト調整（周辺ノードを詰める）
function adjustLayoutAfterCollapse(collapsedNode) {
    console.log('🔄 折りたたみ後のレイアウト調整開始:', collapsedNode.text);
    
    // 表示されているノードのみを対象に衝突解消
    const visibleNodes = getVisibleNodes(root);
    resolveCollisionsForVisibleNodes(visibleNodes);
    
    // アニメーション開始
    startAnimation();
    
    console.log('✅ 折りたたみ後のレイアウト調整完了');
}

// 展開後のレイアウト調整（新たに表示されたノードのスペースを確保）
function adjustLayoutAfterExpansion(expandedNode) {
    console.log('🔄 展開後のレイアウト調整開始:', expandedNode.text);
    
    // まず展開されたノードの子ノードを再配置
    layoutChildrenForNode(expandedNode);
    
    // 表示されているノードのみを対象に衝突解消
    const visibleNodes = getVisibleNodes(root);
    resolveCollisionsForVisibleNodes(visibleNodes);
    
    // アニメーション開始
    startAnimation();
    
    console.log('✅ 展開後のレイアウト調整完了');
}

// 特定ノードの子ノードのみを再配置
function layoutChildrenForNode(parentNode) {
    if (!parentNode.children || parentNode.children.length === 0) return;
    
    console.log('🔄 子ノード再配置:', parentNode.text, '子ノード数:', parentNode.children.length);
    
    // 親ノードの位置を取得
    initializeNodeAnimation(parentNode);
    const parentX = parentNode.targetX || parentNode.x;
    const parentY = parentNode.targetY || parentNode.y;
    
    // レイアウトタイプに応じて子ノードを配置
    parentNode.children.forEach((child, index) => {
        initializeNodeAnimation(child);
        
        let newX, newY;
        
        switch (currentLayout) {
            case 'radial':
                const angle = (index * 60) * Math.PI / 180;
                newX = parentX + Math.cos(angle) * 150;
                newY = parentY + Math.sin(angle) * 150;
                break;
                
            case 'leftright':
                if (index % 2 === 0) {
                    newX = parentX + 180;
                    newY = parentY + (index - parentNode.children.length / 2) * 80;
                } else {
                    newX = parentX - 180;
                    newY = parentY + (index - parentNode.children.length / 2) * 80;
                }
                break;
                
            case 'tree':
            default:
                newX = parentX + 180;
                newY = parentY + (index - (parentNode.children.length - 1) / 2) * 60;
                break;
        }
        
        setNodeTarget(child, newX, newY);
        
        // 子ノードが非折りたたみ状態の場合、さらにその子ノードを再帰配置
        if (!child.collapsed) {
            layoutChildrenForNode(child);
        }
    });
}

// 表示されているノードのみの衝突解消
function resolveCollisionsForVisibleNodes(visibleNodes) {
    console.log('🔄 表示ノードのみで衝突解消開始:', visibleNodes.length, '個のノード');
    
    const maxIterations = 15;
    let hasAdjustment = false;
    
    for (let iteration = 0; iteration < maxIterations; iteration++) {
        let foundCollision = false;
        
        // 表示されているノード同士でのみ衝突チェック
        for (let i = 0; i < visibleNodes.length; i++) {
            for (let j = i + 1; j < visibleNodes.length; j++) {
                const nodeA = visibleNodes[i];
                const nodeB = visibleNodes[j];
                if (checkCollisionByTargetPosition(nodeA, nodeB)) {
                    resolveCollisionWithAnimation(nodeA, nodeB);
                    foundCollision = true;
                    hasAdjustment = true;
                }
            }
        }
        
        // 衝突がなくなったら終了
        if (!foundCollision) break;
    }
    
    if (hasAdjustment) {
        console.log('✅ 表示ノードの衝突解消完了');
    } else {
        console.log('ℹ️ 衝突なし - 調整不要');
    }
}

// Y座標でノードを並び替える関数（再帰版）
function sortNodesByYPosition(node) {
    if (!node || !node.children || node.children.length <= 1) return;
    
    console.log('📋 Y座標ソート開始:', node.text, '子ノード数:', node.children.length);
    
    // 子ノードをY座標順で並び替え
    const beforeSort = node.children.map(child => `${child.text}(y:${child.y})`);
    
    node.children.sort((a, b) => {
        // Y座標でソート（targetYがある場合はそれを優先）
        const yA = a.targetY !== undefined ? a.targetY : a.y;
        const yB = b.targetY !== undefined ? b.targetY : b.y;
        return yA - yB;
    });
    
    const afterSort = node.children.map(child => `${child.text}(y:${child.y})`);
    
    // 順序が変わった場合はログ出力
    if (JSON.stringify(beforeSort) !== JSON.stringify(afterSort)) {
        console.log('✅ Y座標ソート完了:', node.text);
        console.log('  ソート前:', beforeSort.join(' -> '));
        console.log('  ソート後:', afterSort.join(' -> '));
    }
    
    // 再帰的に子ノードもソート
    node.children.forEach(child => {
        sortNodesByYPosition(child);
    });
}

// 全体のノード構造をY座標順で整理する関数
function reorganizeNodesByYPosition() {
    console.log('📋 全体のY座標ソートを実行中...');
    sortNodesByYPosition(root);
    console.log('✅ 全体のY座標ソート完了');
    
    // データの保存
    pushHistory();
    saveCurrentMindMap();
}
function moveNodeUp(node) {
    if (!node.parent) return;
    
    const siblings = node.parent.children;
    const index = siblings.indexOf(node);
    if (index > 0) {
        [siblings[index - 1], siblings[index]] = [siblings[index], siblings[index - 1]];
        pushHistory();
        // 親ノードの子ノードのみ再配置
        setTimeout(() => {
            adjustLayoutAfterDeletion(node.parent);
            // 配置後にY座標順で整理
            setTimeout(() => reorganizeNodesByYPosition(), 100);
        }, 50);
    }
}
function moveNodeDown(node) {
    if (!node.parent) return;
    
    const siblings = node.parent.children;
    const index = siblings.indexOf(node);
    if (index < siblings.length - 1) {
        [siblings[index + 1], siblings[index]] = [siblings[index], siblings[index + 1]];
        pushHistory();
        // 親ノードの子ノードのみ再配置
        setTimeout(() => {
            adjustLayoutAfterDeletion(node.parent);
            // 配置後にY座標順で整理
            setTimeout(() => reorganizeNodesByYPosition(), 100);
        }, 50);
    }
}
// ------------------------------
// レイアウト計算
// ------------------------------
function layout(node = root) {
    if (node === root) {
        // ルートノードを中心に配置（慣性アニメーション付き）
        if (node.x === 0 && node.y === 0) {
            // 初回のみ即座移動
            setNodeImmediately(node, 800, 450);
        } else {
            // 2回目以降は慣性で移動
            setNodeTarget(node, 800, 450);
        }
        
        // レイアウトタイプに応じて子ノードを配置
        switch (currentLayout) {
            case 'radial':
                layoutChildrenRadialWithAnimation(node);
                break;
            case 'leftright':
                layoutChildrenLeftRightWithAnimation(node);
                break;
            case 'tree':
                layoutChildrenTreeWithAnimation(node);
                break;
        }
        
        // アニメーション開始
        startAnimation();
    }
}
// 放射状レイアウト（アニメーション対応）
function layoutChildrenRadialWithAnimation(parent, parentAngle = 0, angleRange = 360) {
    if (parent.children.length === 0) return;
    
    const radius = 150; // 半径
    const angleStep = angleRange / parent.children.length;
    const startAngle = parentAngle - (angleRange / 2) + (angleStep / 2);
    
    parent.children.forEach((child, index) => {
        const angle = startAngle + (index * angleStep);
        const radian = (angle * Math.PI) / 180;
        
        const targetX = parent.targetX + Math.cos(radian) * radius;
        const targetY = parent.targetY + Math.sin(radian) * radius;
        
        setNodeTarget(child, targetX, targetY);
        
        // 子ノードも再帰的にレイアウト
        layoutChildrenRadialWithAnimation(child, angle, 180);
    });
}

// 左右分岐レイアウト（アニメーション対応）
function layoutChildrenLeftRightWithAnimation(parent) {
    if (parent.children.length === 0) return;
    
    const leftChildren = [];
    const rightChildren = [];
    
    // 子ノードを左右に振り分け
    parent.children.forEach((child, index) => {
        if (index % 2 === 0) {
            rightChildren.push(child);
        } else {
            leftChildren.push(child);
        }
    });
    
    layoutChildrenLeftRightRecursiveWithAnimation(parent, 'left', leftChildren);
    layoutChildrenLeftRightRecursiveWithAnimation(parent, 'right', rightChildren);
}

function layoutChildrenLeftRightRecursiveWithAnimation(parent, side, children) {
    const xOffset = side === 'left' ? -180 : 180;
    const ySpacing = 80;
    const startY = parent.targetY - ((children.length - 1) * ySpacing) / 2;
    
    children.forEach((child, index) => {
        const targetX = parent.targetX + xOffset;
        const targetY = startY + (index * ySpacing);
        
        setNodeTarget(child, targetX, targetY);
        
        // 子ノードを再帰的に配置
        layoutChildrenLeftRightWithAnimation(child);
    });
}

// ツリーレイアウト（アニメーション対応）
function layoutChildrenTreeWithAnimation(parent) {
    if (parent.children.length === 0) return;
    
    const xOffset = 180;
    const ySpacing = 60;
    
    layoutChildrenTreeRecursiveWithAnimation(parent, xOffset, ySpacing);
}

function layoutChildrenTreeRecursiveWithAnimation(parent) {
    if (parent.children.length === 0) return;
    
    const xOffset = 180;
    const ySpacing = 60;
    const startY = parent.targetY - ((parent.children.length - 1) * ySpacing) / 2;
    
    parent.children.forEach((child, index) => {
        const targetX = parent.targetX + xOffset;
        const targetY = startY + (index * ySpacing);
        
        setNodeTarget(child, targetX, targetY);
        
        // 子ノードを再帰的に配置
        layoutChildrenTreeRecursiveWithAnimation(child);
    });
}

// 放射状レイアウト（既存）
function layoutChildrenRadial(parent, parentAngle = 0, angleRange = 360) {
    const children = parent.children;
    if (children.length === 0)
        return;
    // 子ノードの数に応じて角度を分割
    const angleStep = angleRange / Math.max(children.length, 1);
    const startAngle = parentAngle - (angleRange / 2) + (angleStep / 2);
    // 階層に応じた距離
    const distance = getDistanceForDepth(getNodeDepth(parent)) + 20;
    children.forEach((child, index) => {
        const angle = startAngle + (angleStep * index);
        const radians = (angle * Math.PI) / 180;
        // 親を中心とした円周上に配置
        child.x = parent.x + Math.cos(radians) * distance;
        child.y = parent.y + Math.sin(radians) * distance;
        // 子ノードがある場合、さらに細かい角度範囲で配置
        if (child.children.length > 0) {
            const childAngleRange = Math.min(angleStep * 0.8, 120); // 最大120度
            layoutChildrenRadial(child, angle, childAngleRange);
        }
    });
}
// 左右分岐レイアウト
function layoutChildrenLeftRight(parent) {
    const children = parent.children;
    if (children.length === 0)
        return;
    const baseDistance = 200;
    const verticalSpacing = 120;
    // 左右に交互に配置
    const leftChildren = children.filter((_, index) => index % 2 === 0);
    const rightChildren = children.filter((_, index) => index % 2 === 1);
    // 左側の子ノード
    leftChildren.forEach((child, index) => {
        child.x = parent.x - baseDistance - (getNodeDepth(child) * 150);
        child.y = parent.y + (index - (leftChildren.length - 1) / 2) * verticalSpacing;
        layoutChildrenLeftRightRecursive(child, 'left');
    });
    // 右側の子ノード
    rightChildren.forEach((child, index) => {
        child.x = parent.x + baseDistance + (getNodeDepth(child) * 150);
        child.y = parent.y + (index - (rightChildren.length - 1) / 2) * verticalSpacing;
        layoutChildrenLeftRightRecursive(child, 'right');
    });
}
// 左右分岐レイアウト（再帰）
function layoutChildrenLeftRightRecursive(parent, side) {
    const children = parent.children;
    if (children.length === 0)
        return;
    const baseDistance = 180;
    const verticalSpacing = 100;
    children.forEach((child, index) => {
        const direction = side === 'left' ? -1 : 1;
        child.x = parent.x + direction * baseDistance;
        child.y = parent.y + (index - (children.length - 1) / 2) * verticalSpacing;
        layoutChildrenLeftRightRecursive(child, side);
    });
}
// ツリーレイアウト
function layoutChildrenTree(parent) {
    const children = parent.children;
    if (children.length === 0)
        return;
    const baseDistance = 200;
    const verticalSpacing = 80;
    // 全て右側に配置
    children.forEach((child, index) => {
        child.x = parent.x + baseDistance;
        child.y = parent.y + (index - (children.length - 1) / 2) * verticalSpacing;
        layoutChildrenTreeRecursive(child);
    });
}
// ツリーレイアウト（再帰）
function layoutChildrenTreeRecursive(parent) {
    const children = parent.children;
    if (children.length === 0)
        return;
    const baseDistance = 180;
    const verticalSpacing = 70;
    children.forEach((child, index) => {
        child.x = parent.x + baseDistance;
        child.y = parent.y + (index - (children.length - 1) / 2) * verticalSpacing;
        layoutChildrenTreeRecursive(child);
    });
}
function layoutChildren(parent, parentAngle = 0, angleRange = 360) {
    // 旧関数は放射状レイアウトを呼び出すように変更
    return layoutChildrenRadial(parent, parentAngle, angleRange);
}
function getNodeDepth(node) {
    if (!node) {
        console.warn('getNodeDepth: nodeが無効です');
        return 0;
    }
    let depth = 0;
    let current = node;
    while (current && current.parent !== null) {
        depth++;
        current = current.parent;
        // 無限ループ防止（循環参照対策）
        if (depth > 100) {
            console.warn('getNodeDepth: 異常な深度が検出されました', depth);
            break;
        }
    }
    return depth;
}
function getDistanceForDepth(depth) {
    // 階層に応じた距離を返す（ルートからの距離）
    // 全体が画面に収まるように短く調整
    const baseDistance = 120; // 180 から 120 に縮小
    const increment = 100; // 140 から 100 に縮小
    return baseDistance + (depth * increment);
}
// ------------------------------
// ノード衝突検出・解消
// ------------------------------
// 慣性アニメーション付き衝突解消
function resolveAllCollisionsWithAnimation() {
    let hasCollisionAdjustment = false;
    
    // まずサブツリー同士の衝突を解決
    hasCollisionAdjustment = resolveSubtreeCollisionsWithAnimation() || hasCollisionAdjustment;
    
    // 次に個別ノードの衝突を解決
    const allNodes = getAllNodes(root);
    const maxIterations = 20;
    
    for (let iteration = 0; iteration < maxIterations; iteration++) {
        let foundCollision = false;
        
        // すべてのノードペアをチェック
        for (let i = 0; i < allNodes.length; i++) {
            for (let j = i + 1; j < allNodes.length; j++) {
                const nodeA = allNodes[i];
                const nodeB = allNodes[j];
                if (checkCollisionByTargetPosition(nodeA, nodeB)) {
                    resolveCollisionWithAnimation(nodeA, nodeB);
                    foundCollision = true;
                    hasCollisionAdjustment = true;
                }
            }
        }
        
        // 衝突がなくなったら終了
        if (!foundCollision) break;
    }
    
    // 衝突解消の調整があった場合はアニメーション開始
    if (hasCollisionAdjustment) {
        console.log('⚡ 衝突解消アニメーション開始');
        startAnimation();
    } else {
        console.log('✅ 衝突なし - アニメーション完全終了');
    }
}

// 目標位置ベースの衝突チェック
function checkCollisionByTargetPosition(nodeA, nodeB) {
    const minDistance = 80; // 最小距離
    const dx = nodeA.targetX - nodeB.targetX;
    const dy = nodeA.targetY - nodeB.targetY;
    const distance = Math.sqrt(dx * dx + dy * dy);
    return distance < minDistance;
}

// アニメーション付き衝突解決
function resolveCollisionWithAnimation(nodeA, nodeB) {
    const minDistance = 80;
    const dx = nodeA.targetX - nodeB.targetX;
    const dy = nodeA.targetY - nodeB.targetY;
    const distance = Math.sqrt(dx * dx + dy * dy);
    
    if (distance < minDistance && distance > 0) {
        const overlap = minDistance - distance;
        const moveDistance = overlap / 2 + 10; // 少し余裕を持たせる
        
        // 正規化された方向ベクトル
        const normalX = dx / distance;
        const normalY = dy / distance;
        
        // 目標位置を調整（現在位置ではなく目標位置を変更）
        nodeA.targetX += normalX * moveDistance;
        nodeA.targetY += normalY * moveDistance;
        nodeB.targetX -= normalX * moveDistance;
        nodeB.targetY -= normalY * moveDistance;
    }
}

// サブツリー衝突解消（アニメーション付き）
function resolveSubtreeCollisionsWithAnimation() {
    if (root.children.length === 0) return false;
    
    let hasAdjustment = false;
    const maxIterations = 15;
    
    for (let iteration = 0; iteration < maxIterations; iteration++) {
        let foundCollision = false;
        
        for (let i = 0; i < root.children.length; i++) {
            for (let j = i + 1; j < root.children.length; j++) {
                const subtreeA = root.children[i];
                const subtreeB = root.children[j];
                if (checkSubtreeCollisionByTarget(subtreeA, subtreeB)) {
                    resolveSubtreeCollisionWithAnimation(subtreeA, subtreeB);
                    foundCollision = true;
                    hasAdjustment = true;
                }
            }
        }
        
        if (!foundCollision) break;
    }
    
    return hasAdjustment;
}

// サブツリー衝突チェック（目標位置ベース）
function checkSubtreeCollisionByTarget(subtreeA, subtreeB) {
    const nodesA = getAllNodes(subtreeA);
    const nodesB = getAllNodes(subtreeB);
    
    for (const nodeA of nodesA) {
        for (const nodeB of nodesB) {
            if (checkCollisionByTargetPosition(nodeA, nodeB)) {
                return true;
            }
        }
    }
    return false;
}

// サブツリー衝突解決（アニメーション付き）
function resolveSubtreeCollisionWithAnimation(subtreeA, subtreeB) {
    const boundsA = getSubtreeBoundsByTarget(subtreeA);
    const boundsB = getSubtreeBoundsByTarget(subtreeB);
    
    const overlapX = Math.max(0, Math.min(boundsA.right, boundsB.right) - Math.max(boundsA.left, boundsB.left));
    const overlapY = Math.max(0, Math.min(boundsA.bottom, boundsB.bottom) - Math.max(boundsA.top, boundsB.top));
    
    if (overlapX > 0 && overlapY > 0) {
        const margin = 20;
        
        if (overlapX < overlapY) {
            // 水平方向に分離
            const moveDistance = (overlapX / 2) + margin;
            
            if (boundsA.centerX < boundsB.centerX) {
                moveSubtreeWithAnimation(subtreeA, -moveDistance, 0);
                moveSubtreeWithAnimation(subtreeB, moveDistance, 0);
            } else {
                moveSubtreeWithAnimation(subtreeA, moveDistance, 0);
                moveSubtreeWithAnimation(subtreeB, -moveDistance, 0);
            }
        } else {
            // 垂直方向に分離
            const moveDistance = (overlapY / 2) + margin;
            
            if (boundsA.centerY < boundsB.centerY) {
                moveSubtreeWithAnimation(subtreeA, 0, -moveDistance);
                moveSubtreeWithAnimation(subtreeB, 0, moveDistance);
            } else {
                moveSubtreeWithAnimation(subtreeA, 0, moveDistance);
                moveSubtreeWithAnimation(subtreeB, 0, -moveDistance);
            }
        }
    }
}

// サブツリーをアニメーションで移動
function moveSubtreeWithAnimation(rootNode, deltaX, deltaY) {
    function moveNodeAndChildren(node) {
        initializeNodeAnimation(node);
        node.targetX += deltaX;
        node.targetY += deltaY;
        
        if (node.children) {
            node.children.forEach(child => moveNodeAndChildren(child));
        }
    }
    
    moveNodeAndChildren(rootNode);
}

// 目標位置ベースのサブツリー範囲取得
function getSubtreeBoundsByTarget(node) {
    const nodes = getAllNodes(node);
    
    let left = Infinity, right = -Infinity;
    let top = Infinity, bottom = -Infinity;
    
    for (const n of nodes) {
        initializeNodeAnimation(n);
        const margin = 40;
        left = Math.min(left, n.targetX - margin);
        right = Math.max(right, n.targetX + margin);
        top = Math.min(top, n.targetY - margin);
        bottom = Math.max(bottom, n.targetY + margin);
    }
    
    return {
        left, right, top, bottom,
        centerX: (left + right) / 2,
        centerY: (top + bottom) / 2
    };
}

function resolveAllCollisions() {
    // まずサブツリー同士の衝突を解決
    resolveSubtreeCollisions();
    // 次に個別ノードの衝突を解決
    const allNodes = getAllNodes(root);
    const maxIterations = 20; // 無限ループを防ぐため
    for (let iteration = 0; iteration < maxIterations; iteration++) {
        let foundCollision = false;
        // すべてのノードペアをチェック
        for (let i = 0; i < allNodes.length; i++) {
            for (let j = i + 1; j < allNodes.length; j++) {
                const nodeA = allNodes[i];
                const nodeB = allNodes[j];
                if (checkCollision(nodeA, nodeB)) {
                    resolveCollision(nodeA, nodeB);
                    foundCollision = true;
                }
            }
        }
        // 衝突がなくなったら終了
        if (!foundCollision)
            break;
    }
}
// サブツリー単位での衝突検出・解決
function resolveSubtreeCollisions() {
    if (root.children.length === 0)
        return;
    const maxIterations = 15; // サブツリー衝突解決の最大繰り返し回数
    for (let iteration = 0; iteration < maxIterations; iteration++) {
        let foundCollision = false;
        // ルートの全子ノード（メインブランチ）同士をチェック
        for (let i = 0; i < root.children.length; i++) {
            for (let j = i + 1; j < root.children.length; j++) {
                const subtreeA = root.children[i];
                const subtreeB = root.children[j];
                if (checkSubtreeCollision(subtreeA, subtreeB)) {
                    resolveSubtreeCollision(subtreeA, subtreeB);
                    foundCollision = true;
                }
            }
        }
        if (!foundCollision)
            break;
    }
}
function getSubtreeBounds(rootNode) {
    const allNodes = getAllNodesInSubtree(rootNode);
    if (allNodes.length === 0) {
        const nodeWidth = Math.max(120, rootNode.text.length * 8) / 2;
        const nodeHeight = 20;
        return {
            minX: rootNode.x - nodeWidth,
            maxX: rootNode.x + nodeWidth,
            minY: rootNode.y - nodeHeight,
            maxY: rootNode.y + nodeHeight
        };
    }
    // 最初のノードの動的サイズで初期化
    const firstNodeWidth = Math.max(120, allNodes[0].text.length * 8) / 2;
    let minX = allNodes[0].x - firstNodeWidth;
    let maxX = allNodes[0].x + firstNodeWidth;
    let minY = allNodes[0].y - 20; // ノード高さの半分
    let maxY = allNodes[0].y + 20;
    allNodes.forEach(node => {
        const nodeWidth = Math.max(120, node.text.length * 8) / 2; // 動的幅の半分
        minX = Math.min(minX, node.x - nodeWidth);
        maxX = Math.max(maxX, node.x + nodeWidth);
        minY = Math.min(minY, node.y - 20);
        maxY = Math.max(maxY, node.y + 20);
    });
    return { minX, maxX, minY, maxY };
}
// サブツリー内の全ノードを取得
function getAllNodesInSubtree(rootNode) {
    const result = [rootNode];
    function traverse(node) {
        for (const child of node.children) {
            result.push(child);
            traverse(child);
        }
    }
    traverse(rootNode);
    return result;
}
// サブツリー同士の衝突チェック
function checkSubtreeCollision(subtreeA, subtreeB) {
    const boundsA = getSubtreeBounds(subtreeA);
    const boundsB = getSubtreeBounds(subtreeB);
    // マージンを追加して余裕を持たせる
    const margin = 40;
    // 境界ボックスの重複チェック
    const noOverlapX = boundsA.maxX + margin < boundsB.minX || boundsB.maxX + margin < boundsA.minX;
    const noOverlapY = boundsA.maxY + margin < boundsB.minY || boundsB.maxY + margin < boundsA.minY;
    return !(noOverlapX || noOverlapY);
}
// サブツリー衝突の解決
function resolveSubtreeCollision(subtreeA, subtreeB) {
    const boundsA = getSubtreeBounds(subtreeA);
    const boundsB = getSubtreeBounds(subtreeB);
    // 重複している領域を計算
    const overlapX = Math.min(boundsA.maxX, boundsB.maxX) - Math.max(boundsA.minX, boundsB.minX);
    const overlapY = Math.min(boundsA.maxY, boundsB.maxY) - Math.max(boundsA.minY, boundsB.minY);
    // より小さい重複方向に移動
    const margin = 50; // 追加の分離マージン
    if (overlapX < overlapY) {
        // X方向に分離
        const centerA = (boundsA.minX + boundsA.maxX) / 2;
        const centerB = (boundsB.minX + boundsB.maxX) / 2;
        const moveDistance = (overlapX / 2) + margin;
        if (centerA < centerB) {
            // AをX負方向、BをX正方向に移動
            moveSubtree(subtreeA, -moveDistance, 0);
            moveSubtree(subtreeB, moveDistance, 0);
        }
        else {
            // AをX正方向、BをX負方向に移動
            moveSubtree(subtreeA, moveDistance, 0);
            moveSubtree(subtreeB, -moveDistance, 0);
        }
    }
    else {
        // Y方向に分離
        const centerA = (boundsA.minY + boundsA.maxY) / 2;
        const centerB = (boundsB.minY + boundsB.maxY) / 2;
        const moveDistance = (overlapY / 2) + margin;
        if (centerA < centerB) {
            // AをY負方向、BをY正方向に移動
            moveSubtree(subtreeA, 0, -moveDistance);
            moveSubtree(subtreeB, 0, moveDistance);
        }
        else {
            // AをY正方向、BをY負方向に移動
            moveSubtree(subtreeA, 0, moveDistance);
            moveSubtree(subtreeB, 0, -moveDistance);
        }
    }
}
// サブツリー全体を移動
function moveSubtree(rootNode, deltaX, deltaY) {
    const allNodes = getAllNodesInSubtree(rootNode);
    allNodes.forEach(node => {
        node.x += deltaX;
        node.y += deltaY;
    });
}
function checkCollision(nodeA, nodeB) {
    const dx = nodeB.x - nodeA.x;
    const dy = nodeB.y - nodeA.y;
    const distance = Math.sqrt(dx * dx + dy * dy);
    const minDistance = getMinDistanceBetweenNodes(nodeA, nodeB);
    return distance < minDistance;
}
function resolveCollision(nodeA, nodeB) {
    const dx = nodeB.x - nodeA.x;
    const dy = nodeB.y - nodeA.y;
    const distance = Math.sqrt(dx * dx + dy * dy);
    const minDistance = getMinDistanceBetweenNodes(nodeA, nodeB);
    if (distance === 0) {
        // 完全に重複している場合のランダム移動
        const angle = Math.random() * Math.PI * 2;
        const offset = minDistance / 2;
        nodeA.x -= Math.cos(angle) * offset;
        nodeA.y -= Math.sin(angle) * offset;
        nodeB.x += Math.cos(angle) * offset;
        nodeB.y += Math.sin(angle) * offset;
        return;
    }
    // 正規化された方向ベクトル
    const normalX = dx / distance;
    const normalY = dy / distance;
    // 必要な移動距離
    const overlap = minDistance - distance;
    // ルートノードは動かさない
    if (nodeA === root) {
        nodeB.x += normalX * overlap;
        nodeB.y += normalY * overlap;
    }
    else if (nodeB === root) {
        nodeA.x -= normalX * overlap;
        nodeA.y -= normalY * overlap;
    }
    else {
        // どちらも移動可能な場合、階層の深い方を優先的に移動
        const depthA = getNodeDepth(nodeA);
        const depthB = getNodeDepth(nodeB);
        let moveRatioA = 0.5;
        let moveRatioB = 0.5;
        if (depthA > depthB) {
            // nodeAの方が深い階層 → nodeAを多く移動
            moveRatioA = 0.7;
            moveRatioB = 0.3;
        }
        else if (depthB > depthA) {
            // nodeBの方が深い階層 → nodeBを多く移動
            moveRatioA = 0.3;
            moveRatioB = 0.7;
        }
        nodeA.x -= normalX * overlap * moveRatioA;
        nodeA.y -= normalY * overlap * moveRatioA;
        nodeB.x += normalX * overlap * moveRatioB;
        nodeB.y += normalY * overlap * moveRatioB;
    }
}
function getMinDistanceBetweenNodes(nodeA, nodeB) {
    // ノードサイズ考慮した基本距離
    const baseDistance = 140;
    // 親子関係の場合は少し近くても良い
    if (isParentChild(nodeA, nodeB)) {
        return baseDistance * 0.9; // 10%近く
    }
    // 兄弟関係の場合は標準距離
    if (areSiblings(nodeA, nodeB)) {
        return baseDistance;
    }
    // その他の関係は少し遠く
    return baseDistance * 1.1;
}
function isParentChild(nodeA, nodeB) {
    return nodeA.parent === nodeB || nodeB.parent === nodeA;
}
function areSiblings(nodeA, nodeB) {
    // 安全性チェック
    if (!nodeA || !nodeB || !nodeA.parent || !nodeB.parent) {
        return false;
    }
    return nodeA.parent === nodeB.parent;
}
function getMinNodeDistance() {
    // ノード間の最小距離（ノードサイズ考慮）
    // ノードの幅120px + マージン20px = 140px
    return 140;
}
// スムーズなアニメーション用のレイアウト更新
function animatedLayout() {
    const oldPositions = new Map();
    // 現在の位置を記録
    function recordPositions(node) {
        oldPositions.set(node.id, { x: node.x, y: node.y });
        node.children.forEach(recordPositions);
    }
    recordPositions(root);
    // 新しいレイアウトを計算
    layout(root);
    // アニメーションで移動
    animateToNewPositions(root, oldPositions);
}
function animateToNewPositions(node, oldPositions) {
    const old = oldPositions.get(node.id);
    if (old) {
        const startTime = Date.now();
        const duration = 500; // 500ms のアニメーション
        const startX = old.x;
        const startY = old.y;
        const endX = node.x;
        const endY = node.y;
        function animate() {
            const elapsed = Date.now() - startTime;
            const progress = Math.min(elapsed / duration, 1);
            // イージング関数（スムーズな動き）
            const eased = 1 - Math.pow(1 - progress, 3);
            node.x = startX + (endX - startX) * eased;
            node.y = startY + (endY - startY) * eased;
            if (progress < 1) {
                requestAnimationFrame(animate);
            }
            draw();
        }
        animate();
    }
    // 子ノードも同様にアニメーション
    node.children.forEach(child => animateToNewPositions(child, oldPositions));
}
// ------------------------------
// 描画
// ------------------------------
function draw() {
    svg.innerHTML = "";
    // アニメーション中でない初回または通常レイアウト
    if (!isAnimating && (root.children.length === 0 || root.x === 0)) {
        layout(root);
    }
    drawConnections(root);
    drawNodes(root);
}
function drawNodes(node) {
    drawNode(node);
    // 折りたたまれていない場合のみ子ノードを描画
    if (!node.collapsed) {
        for (const child of node.children)
            drawNodes(child);
    }
}
function drawNode(node) {
    const group = document.createElementNS("http://www.w3.org/2000/svg", "g");
    // ノードIDを属性として設定（編集時の要素特定に使用）
    group.setAttribute("data-node-id", node.id);
    
    // パンオフセットとズームを適用した座標
    const x = (node.x + panX) * zoom;
    const y = (node.y + panY) * zoom;
    // フォントサイズをズームに完全連動（最小値制限なし）
    const fontSize = 12 * zoom; // ズームに完全連動
    // テキストの長さに応じてノード幅を調整
    const textLength = node.text.length;
    const baseWidth = Math.max(120, textLength * 8); // 最低120px、文字あたり8px
    const width = baseWidth * zoom;
    const height = 40 * zoom;
    // 極小ズーム時はテキストを簡略化
    let displayText = node.text;
    if (zoom < 0.2) {
        displayText = textLength > 4 ? node.text.substring(0, 3) + '.' : node.text;
    }
    else if (zoom < 0.4) {
        displayText = textLength > 8 ? node.text.substring(0, 6) + '..' : node.text;
    }
    else if (zoom < 0.7) {
        displayText = textLength > 12 ? node.text.substring(0, 10) + '...' : node.text;
    }
    const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.setAttribute("x", String(x - width / 2));
    rect.setAttribute("y", String(y - height / 2));
    rect.setAttribute("width", String(width));
    rect.setAttribute("height", String(height));
    rect.setAttribute("rx", String(8 * zoom));
    // 選択状態でスタイルを強調し、アニメーション効果を追加
    const isSelected = node === getSelected();
    const isDragTarget = node === draggingNode && isDragging;
    const isCollapsed = node.collapsed && node.children && node.children.length > 0;
    // 基本スタイル
    let fillColor = "#fff";
    let strokeColor = "#333";
    let strokeWidth = Math.max(1, zoom);
    let textColor = "#333";
    let fontWeight = "400";
    let strokeDasharray = "none";

    // トップノード判定（root.idと一致するノードのみ）
    const isRoot = (typeof root !== 'undefined') && node.id === root.id;
    // 第一階層ノード判定（root.childrenのみ）
    const isFirstLayer = (typeof root !== 'undefined') && root.children && root.children.some(child => child.id === node.id);

    // トップノード強調（濃い青系）
    if (isRoot) {
        fillColor = "#174378"; // 濃い青
        strokeColor = "#0d2544"; // さらに濃い青
        strokeWidth = Math.max(4, zoom * 2);
        textColor = "#fff";
        fontWeight = "bold";
    } else if (isFirstLayer) {
        // 第一階層ノード強調（さらに薄い青）
        fillColor = "#e3eaf6";
        strokeColor = "#174378";
        strokeWidth = Math.max(3, zoom * 1.5);
    }

    // 第二階層ノード判定（選択状態より優先度低）
    const isSecondLayer = typeof root !== 'undefined' && node.parent && root.children && root.children.some(child => child.id === node.parent.id);
    
    // isDoneとisAccentの事前判定
    const isDone = node.text && node.text.includes('★');  // ★ = 完了・重要（グレー系）
    const isAccent = node.text && node.text.includes('☆');  // ☆ = 強調・アクセント（黄色系）
    
    // ☆ノード（黄色系で発光・強調）- 最優先状態以外の場合に適用
    if (isAccent && !isSelected && !isDragTarget && !isCollapsed) {
        fillColor = "#fffbe6";
        strokeColor = "#ffc107";
        strokeWidth = Math.max(3, zoom * 2);
        textColor = "#d48806";
        fontWeight = "bold";
        rect && rect.setAttribute && rect.setAttribute("filter", `drop-shadow(0 0 12px #ffd70088)`);
    }
    
    // ★ノード（グレー系で発光・完了）- ☆がない場合に適用
    else if (isDone && !isAccent && !isSelected && !isDragTarget && !isCollapsed) {
        fillColor = "#f3f3f3";
        strokeColor = "#9e9e9e";
        strokeWidth = Math.max(3, zoom * 2);
        textColor = "#5d4037";
        fontWeight = "bold";
        rect && rect.setAttribute && rect.setAttribute("filter", `drop-shadow(0 0 12px #9e9e9e88)`);
    }
    
    // 第二階層ノードの基本スタイル（isDone、isAccent、その他優先状態以外の場合のみ適用）
    else if (isSecondLayer && !isSelected && !isDragTarget && !isCollapsed) {
        fillColor = "#f5f8fc";
    }
    
    // 以下の状態は最優先（上記のスタイルを上書き）
    if (isCollapsed) {
        fillColor = "#f0f8ff";
        strokeColor = "#4682b4";
        strokeWidth = Math.max(2, zoom * 2);
        strokeDasharray = "5,3";
        textColor = "#2e4a6b";
        fontWeight = "500";
    } else if (isDragTarget) {
        fillColor = "#fff3e0";
        strokeColor = "#ff9800";
        strokeWidth = Math.max(3, zoom * 3);
        textColor = "#e65100";
        fontWeight = "600";
    } else if (isSelected) {
        fillColor = "#e3f2fd";
        strokeColor = "#1976d2";
        strokeWidth = Math.max(2, zoom * 2);
        textColor = "#1976d2";
        fontWeight = "600";
    }
    rect.setAttribute("fill", fillColor);
    rect.setAttribute("stroke", strokeColor);
    rect.setAttribute("stroke-width", String(strokeWidth));

    // 点線スタイルを適用
    if (strokeDasharray !== "none") {
        rect.setAttribute("stroke-dasharray", strokeDasharray);
    }

    // 選択またはドラッグ中、または☆/★ノードにドロップシャドウ効果
    if (isSelected || isDragTarget) {
        rect.setAttribute("filter", `drop-shadow(0 ${Math.max(2, zoom * 2)}px ${Math.max(4, zoom * 4)}px rgba(25, 118, 210, 0.3))`);
    } else if (isDone) {
        rect.setAttribute("filter", `drop-shadow(0 0 12px #99989888)`);
    } else if (isAccent) {
        rect.setAttribute("filter", `drop-shadow(0 0 12px #ffd70088)`);
    }

        // --- 文字色・太さの最終上書き ---
        // 第一階層または第二階層ノードなら必ず青・太字（どんな状態でも最終的に上書き）
        if (
            (isFirstLayer || (typeof root !== 'undefined' && node.parent && root.children && root.children.some(child => child.id === node.parent.id)))
        ) {
            textColor = "#174378";
            fontWeight = "bold";
        }
        const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
        text.setAttribute("x", String(x));
        text.setAttribute("y", String(y + fontSize / 3));
        text.setAttribute("text-anchor", "middle");
        text.setAttribute("font-size", String(fontSize));
        // ここで一度仮の色・太さをセット
        text.setAttribute("fill", textColor);
        text.setAttribute("font-weight", fontWeight);
        text.textContent = displayText;

        // --- 文字色・太さの最終上書き ---
        // 第一階層は枠も青・太字、第二階層は文字だけ青・太字
        if (isFirstLayer) {
            text.setAttribute("fill", "#174378");
            text.setAttribute("font-weight", "bold");
        } else if (typeof root !== 'undefined' && node.parent && root.children && root.children.some(child => child.id === node.parent.id)) {
            text.setAttribute("fill", "#174378");
            text.setAttribute("font-weight", "bold");
            rect.setAttribute("fill", "#f5f8fc"); // とても薄い青
        }
    // 折りたたみインジケーターを表示
    if (node.children && node.children.length > 0) {
        const indicator = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        const indicatorSize = Math.max(12 * zoom, 8);
        indicator.setAttribute("cx", String(x + width/2 - 10 * zoom));
        indicator.setAttribute("cy", String(y - height/2 + 8 * zoom));
        indicator.setAttribute("r", String(indicatorSize / 2));
        indicator.setAttribute("fill", node.collapsed ? "#ff6b6b" : "#51cf66");
        indicator.setAttribute("stroke", "#fff");
        indicator.setAttribute("stroke-width", "2");
        indicator.style.cursor = "pointer";
        
        // インジケーターのテキスト
        const indicatorText = document.createElementNS("http://www.w3.org/2000/svg", "text");
        indicatorText.setAttribute("x", String(x + width/2 - 10 * zoom));
        indicatorText.setAttribute("y", String(y - height/2 + 8 * zoom + 4));
        indicatorText.setAttribute("text-anchor", "middle");
        indicatorText.setAttribute("font-size", String(Math.max(10 * zoom, 8)));
        indicatorText.setAttribute("fill", "white");
        indicatorText.setAttribute("font-weight", "bold");
        indicatorText.textContent = node.collapsed ? "+" : "-";
        indicatorText.style.cursor = "pointer";
        indicatorText.style.pointerEvents = "none"; // テキストはクリックを通す
        
        // インジケーターのクリックイベント
        indicator.addEventListener("click", (e) => {
            e.preventDefault();
            e.stopPropagation();
            const wasCollapsed = node.collapsed;
            node.collapsed = !node.collapsed;
            console.log('📁 折りたたみインジケータークリック:', node.text, 'collapsed:', node.collapsed);
            
            // 折りたたみ状態の変化に応じてレイアウトを調整
            setTimeout(() => {
                if (wasCollapsed && !node.collapsed) {
                    // 展開時: 子ノードを再表示し、周辺ノードを調整
                    console.log('🔄 インジケータークリックで展開時のレイアウト調整');
                    adjustLayoutAfterExpansion(node);
                } else if (!wasCollapsed && node.collapsed) {
                    // 折りたたみ時: 周辺ノードを詰める
                    console.log('🔄 インジケータークリックで折りたたみ時のレイアウト調整');
                    adjustLayoutAfterCollapse(node);
                }
            }, 50);
            
            pushHistory();
            saveCurrentMindMap();
            draw();
        });
        
        group.appendChild(indicator);
        group.appendChild(indicatorText);
    }
    
    group.appendChild(rect);
    group.appendChild(text);
    // グループ要素にクラスを適用
    if (isSelected) {
        group.classList.add('selected');
    }
    if (isDragTarget) {
        group.classList.add('dragging');
    }
    if (isCollapsed) {
        group.classList.add('collapsed');
    }
    svg.appendChild(group);
    
    // シンプルなクリック検出システム
    let lastClickTime = 0;
    let clickCount = 0;
    
    group.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        const currentTime = Date.now();
        const timeDiff = currentTime - lastClickTime;
        
        if (timeDiff < 400) { // 400ms以内なら連続クリック
            clickCount++;
        } else {
            clickCount = 1; // 時間が空いたので新しいクリック
        }
        
        lastClickTime = currentTime;
        
        console.log(`🖱️ クリック検出: ${clickCount}回目, 間隔: ${timeDiff}ms`);
        
        if (clickCount === 1) {
            // シングルクリック処理（少し遅延）
            setTimeout(() => {
                if (clickCount === 1) { // まだ1回のままなら実行
                    console.log('========== シングルクリック ==========');
                    setSelected(node);
                    
                    if (e.ctrlKey) {
                        console.log('🎯 Ctrl+クリック: ノードを画面中央に移動');
                        centerNodeOnScreen(node);
                    }
                    
                    draw();
                    console.log('シングルクリック完了');
                }
            }, 350);
            
        } else if (clickCount === 2) {
            // ダブルクリック処理（即座に実行）
            console.log('🔥 ========== ダブルクリック ==========');
            
            if (node.children && node.children.length > 0) {
                const wasCollapsed = node.collapsed;
                node.collapsed = !node.collapsed;
                console.log('📁 折りたたみ状態変更:', node.text, 'collapsed:', node.collapsed);
                
                setSelected(node);
                
                // レイアウト調整
                setTimeout(() => {
                    if (wasCollapsed && !node.collapsed) {
                        console.log('🔄 展開時のレイアウト調整');
                        adjustLayoutAfterExpansion(node);
                    } else if (!wasCollapsed && node.collapsed) {
                        console.log('🔄 折りたたみ時のレイアウト調整');
                        adjustLayoutAfterCollapse(node);
                    }
                }, 50);
                
                pushHistory();
                saveCurrentMindMap();
                draw();
            } else {
                console.log('🔄 子ノードがないため折りたたみ不可');
                setSelected(node);
                draw();
            }
            
            console.log('🔥 ダブルクリック完了 ==========');
            
            // カウンターリセット
            setTimeout(() => {
                clickCount = 0;
            }, 100);
        }
    });
    
    group.addEventListener("mousedown", (e) => {
        console.log('👆 mousedown:', node.text);
        
        // 編集中でなければドラッグ準備
        if (!editingNode) {
            e.stopPropagation(); // パンイベントを阻止
            // ドラッグ開始の準備
            draggingNode = node;
            dragOffsetX = e.clientX - (node.x + panX) * zoom;
            dragOffsetY = e.clientY - (node.y + panY) * zoom;
            dragStartX = e.clientX;
            dragStartY = e.clientY;
            isDragging = false;
            clickStartTime = Date.now();
        }
    });
}
function drawConnections(node, depth = 0) {
    // 折りたたまれたノードの子ノードは接続線を描画しない
    if (node.collapsed) {
        return;
    }
    
    for (const child of node.children) {
        // ベジェ曲線でマインドマップらしい線を描画
        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        // パンオフセットとズームを適用した座標
        const x1 = (node.x + panX) * zoom;
        const y1 = (node.y + panY) * zoom;
        const x2 = (child.x + panX) * zoom;
        const y2 = (child.y + panY) * zoom;
        // 制御点を計算（滑らかな曲線のため）
        const dx = x2 - x1;
        const dy = y2 - y1;
        const distance = Math.sqrt(dx * dx + dy * dy);
        // 制御点の位置を調整（距離に応じて曲がり具合を調整）
        const curveFactor = Math.min(distance * 0.4, 100 * zoom);
        const cp1x = x1 + curveFactor;
        const cp1y = y1;
        const cp2x = x2 - curveFactor;
        const cp2y = y2;
        // SVGパスの作成（三次ベジェ曲線）
        const pathData = `M ${x1} ${y1} C ${cp1x} ${cp1y}, ${cp2x} ${cp2y}, ${x2} ${y2}`;
        path.setAttribute("d", pathData);
        path.setAttribute("fill", "none");
        path.setAttribute("stroke-linecap", "round");
        path.setAttribute("opacity", "0.9");

        // トップノード→第一階層ノードだけ太く濃紺色
        let color = "#1976d2";
        let strokeWidth = Math.max(3 - depth * 0.3, 1.5) * zoom;
        // トップノード→第一階層
        if (typeof root !== 'undefined' && node.id === root.id && depth === 0) {
            color = "#102040";
            strokeWidth = 8 * zoom;
            path.setAttribute("filter", "drop-shadow(0 0 8px #10204088)");
        // 第一階層→第二階層
        } else if (typeof root !== 'undefined' && root.children && root.children.some(child => child.id === node.id) && depth === 1) {
            color = "#174378";
            strokeWidth = 4.5 * zoom;
            path.setAttribute("filter", "drop-shadow(0 0 4px #17437888)");
        }
        path.setAttribute("stroke", color);
        path.setAttribute("stroke-width", strokeWidth.toString());

        // 距離による太さ微調整（通常線のみ）
        if (!((typeof root !== 'undefined' && node.id === root.id && depth === 0) || (typeof root !== 'undefined' && root.children && root.children.some(child => child.id === node.id) && depth === 1))) {
            const scaledDistance300 = 300 * zoom;
            const scaledDistance150 = 150 * zoom;
            if (distance > scaledDistance300) {
                path.setAttribute("stroke-width", (strokeWidth + 0.5 * zoom).toString());
            }
            else if (distance < scaledDistance150) {
                path.setAttribute("stroke-width", Math.max(strokeWidth - 0.2 * zoom, 0.5).toString());
            }
        }
        svg.appendChild(path);

        // 折りたたまれていない場合のみ再帰的に子ノードの接続線を描画
        if (!child.collapsed) {
            drawConnections(child, depth + 1);
        }
    }
}
// ------------------------------
// JSON 保存
// ------------------------------
function saveJSON() {
    const data = JSON.stringify(root, replacer, 2);
    const blob = new Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "mindmap.json";
    a.click();
    URL.revokeObjectURL(url);
}
// ------------------------------
// JSON 読み込み
// ------------------------------
function loadJSON() {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "application/json";
    input.onchange = () => {
        var _a;
        const file = (_a = input.files) === null || _a === void 0 ? void 0 : _a[0];
        if (!file)
            return;
        const reader = new FileReader();
        reader.onload = () => {
            const json = reader.result;
            loadSnapshot(json);
            pushHistory();
            saveCurrentMindMap(); // JSON読み込み時に即座保存
        };
        reader.readAsText(file);
    };
    input.click();
}
// ------------------------------
// 視覚的ナビゲーション
// ------------------------------
// 表示されているノードのみを取得する関数（折りたたまれた子ノードは除外）
function getVisibleNodes(node = root) {
    const nodes = [node];
    if (!node.collapsed && node.children) {
        node.children.forEach(child => {
            nodes.push(...getVisibleNodes(child));
        });
    }
    return nodes;
}

// 全ノードを収集する関数
function getAllNodes(node = root) {
    const nodes = [node];
    node.children.forEach(child => {
        nodes.push(...getAllNodes(child));
    });
    return nodes;
}
// 指定した方向で最も近いノードを見つける
function findClosestNodeInDirection(currentNode, direction) {
    // 表示されているノードのみを対象にする
    const allNodes = getVisibleNodes(root);
    const candidateNodes = allNodes.filter(node => node !== currentNode);
    if (candidateNodes.length === 0)
        return null;
    // ズームとパンを適用した画面上の座標で比較
    const currentX = (currentNode.x + panX) * zoom;
    const currentY = (currentNode.y + panY) * zoom;
    // 方向に応じてフィルタリング
    const validNodes = candidateNodes.filter(node => {
        const nodeX = (node.x + panX) * zoom;
        const nodeY = (node.y + panY) * zoom;
        switch (direction) {
            case 'up':
                return nodeY < currentY - 10 * zoom;
            case 'down':
                return nodeY > currentY + 10 * zoom;
            case 'left':
                return nodeX < currentX - 10 * zoom;
            case 'right':
                return nodeX > currentX + 10 * zoom;
            default:
                return false;
        }
    });
    if (validNodes.length === 0)
        return null;
    // 最も近いノードを見つける（方向を重視した距離計算）
    let closestNode = validNodes[0];
    let minDistance = getDirectionalDistance(currentNode, closestNode, direction);
    validNodes.forEach(node => {
        const distance = getDirectionalDistance(currentNode, node, direction);
        if (distance < minDistance) {
            minDistance = distance;
            closestNode = node;
        }
    });
    return closestNode;
}
// 方向を重視した距離計算
function getDirectionalDistance(from, to, direction) {
    // ズームとパンを適用した画面上の座標で計算
    const fromX = (from.x + panX) * zoom;
    const fromY = (from.y + panY) * zoom;
    const toX = (to.x + panX) * zoom;
    const toY = (to.y + panY) * zoom;
    const dx = toX - fromX;
    const dy = toY - fromY;
    // 方向の主軸に重みを付けた距離計算
    switch (direction) {
        case 'up':
        case 'down':
            // 縦方向移動では縦の差を重視
            return Math.abs(dy) + Math.abs(dx) * 0.5;
        case 'left':
        case 'right':
            // 横方向移動では横の差を重視
            return Math.abs(dx) + Math.abs(dy) * 0.5;
        default:
            return Math.sqrt(dx * dx + dy * dy);
    }
}
// ------------------------------
// ドラッグ移動
// ------------------------------
let draggingNode = null;
let dragOffsetX = 0;
let dragOffsetY = 0;
let isDragging = false;
let dragStartX = 0;
let dragStartY = 0;
let clickStartTime = 0;
// ------------------------------
// キャンバスパン
// ------------------------------
let panX = 0;
let panY = 0;
let isPanning = false;
let panStartX = 0;
let panStartY = 0;
// ------------------------------
// ズーム機能
// ------------------------------
let zoom = 1.0;
const minZoom = 0.1;
const maxZoom = 3.0;
// ------------------------------
// インライン編集
// ------------------------------
let editingNode = null;
let editingInput = null;
function startInlineEdit(node, event) {
    // 既に編集中なら終了
    if (editingNode) {
        finishEdit();
    }
    editingNode = node;
    setSelected(node);
    createEditInput(node, node.text);
}
function startInlineEditWithKey(node, initialKey) {
    // 既に編集中なら終了
    if (editingNode) {
        finishEdit();
    }
    editingNode = node;
    setSelected(node);
    // 初期キーで置き換え開始
    createEditInput(node, initialKey);
}
function createEditInput(node, initialText) {
    // 入力フィールドを作成
    const input = document.createElement('input');
    input.type = 'text';
    input.value = initialText;
    input.style.position = 'absolute';
    
    // 実際に描画されているSVG要素から位置を取得する方法を試す
    let screenX, screenY;
    
    // ノードIDを使ってSVG内の実際の要素を探す
    const nodeElement = svg.querySelector(`[data-node-id="${node.id}"] rect`) || 
                       svg.querySelector(`g[data-node-id="${node.id}"]`);
    
    if (nodeElement) {
        // SVG要素の実際の描画位置を取得
        const elementRect = nodeElement.getBoundingClientRect();
        screenX = elementRect.left + elementRect.width / 2;
        screenY = elementRect.top + elementRect.height / 2;
    } else {
        // フォールバック: SVGの座標系から計算
        const svgRect = svg.getBoundingClientRect();
        screenX = svgRect.left + node.x;
        screenY = svgRect.top + node.y;
    }
    
    // 編集ボックスのサイズ
    const inputWidth = 150;
    const inputHeight = 30;
    const nodeHeight = 30; // ノードの高さ
    const margin = 8; // ノードとの間隔
    
    // ノードの下に配置（中央寄せ）
    input.style.left = (screenX - inputWidth / 2) + 'px';
    input.style.top = (screenY + nodeHeight / 2 + margin) + 'px';
    input.style.width = inputWidth + 'px';
    input.style.height = inputHeight + 'px';
    input.style.border = '2px solid #007acc';
    input.style.borderRadius = '4px';
    input.style.textAlign = 'center';
    input.style.fontSize = '14px';
    input.style.zIndex = '1000';
    input.style.backgroundColor = 'white';
    input.style.padding = '4px';
    input.style.boxSizing = 'border-box';
    // SVGの親要素に追加（相対配置対応）
    const container = svg.parentElement;
    if (container) {
        container.appendChild(input);
        editingInput = input;
        input.focus();
        // 初期テキストがノードの既存テキストなら選択、新しいキーなら末尾にカーソル
        if (initialText === node.text) {
            input.select();
        }
        else {
            input.setSelectionRange(initialText.length, initialText.length);
        }
        input.select();
        // エンターキーで確定
        input.addEventListener('keydown', (e) => {
            e.stopPropagation(); // イベントの伝播を停止
            if (e.key === 'Enter') {
                finishEdit();
                e.preventDefault();
            }
            else if (e.key === 'Escape') {
                cancelEdit();
                e.preventDefault();
            }
        });
        // フォーカスが外れたら確定
        input.addEventListener('blur', () => {
            finishEdit();
        });
    }
}
function finishEdit() {
    if (editingNode && editingInput) {
        const newText = editingInput.value.trim();
        if (newText && newText !== editingNode.text) {
            editingNode.text = newText;
            pushHistory();
            saveCurrentMindMap(); // ノード編集完了時に即座保存
        }
        cleanup();
    }
}
function cancelEdit() {
    cleanup();
}
function cleanup() {
    if (editingInput) {
        // 安全にDOM要素を削除
        try {
            // 複数の確認方法で安全性を向上
            if (editingInput.parentNode) {
                editingInput.remove(); // より安全なremoveメソッドを使用
            }
        } catch (error) {
            // エラーが発生しても処理を継続（デバッグ用のログのみ出力）
            console.debug('編集要素の削除時にエラー:', error.message);
        }
    }
    editingInput = null;
    editingNode = null;
    draw();
}
// マウスイベント初期化
// ノード間の距離を計算
function getDistance(node1, node2) {
    const dx = node1.x - node2.x;
    const dy = node1.y - node2.y;
    return Math.sqrt(dx * dx + dy * dy);
}

// 指定したノードと重なっている他のノードを検索
function findOverlappingNode(draggedNode, excludeNode = null, searchRadius = 30) {
    function searchInNode(node) {
        // 自分自身、除外ノード、およびドラッグ中のノードの子ノードはスキップ
        if (node === draggedNode || node === excludeNode || isDescendantOf(node, draggedNode)) {
            return null;
        }
        
        // 距離をチェック
        if (getDistance(draggedNode, node) < searchRadius) {
            return node;
        }
        
        // 子ノードも検索
        for (const child of node.children) {
            const found = searchInNode(child);
            if (found) return found;
        }
        
        return null;
    }
    
    return searchInNode(root);
}

// ノードAがノードBの子孫かどうかをチェック
function isDescendantOf(nodeA, nodeB) {
    function checkInChildren(parent) {
        for (const child of parent.children) {
            if (child === nodeA) return true;
            if (checkInChildren(child)) return true;
        }
        return false;
    }
    return checkInChildren(nodeB);
}

// ツリーから指定ノードを削除
function removeNodeFromTree(nodeToRemove) {
    function removeFromNode(node) {
        const index = node.children.indexOf(nodeToRemove);
        if (index !== -1) {
            node.children.splice(index, 1);
            return true;
        }
        
        for (const child of node.children) {
            if (removeFromNode(child)) return true;
        }
        return false;
    }
    
    return removeFromNode(root);
}

// ノードを新しい親の子として追加
function addNodeAsChild(parentNode, childNode) {
    if (!parentNode.children) {
        parentNode.children = [];
    }
    parentNode.children.push(childNode);
}

// ノードとその配下ノードを一緒に移動させる関数
function moveNodeWithChildren(node, deltaX, deltaY) {
    // メインノードを移動
    node.x += deltaX;
    node.y += deltaY;
    
    // 配下ノードを再帰的に移動
    if (node.children && node.children.length > 0) {
        node.children.forEach(child => {
            moveNodeWithChildren(child, deltaX, deltaY);
        });
    }
}

function initializeMouseEvents() {
    console.log('マウスイベント初期化開始');
    if (!svg) {
        console.error('SVG要素が見つからないため、マウスイベントを初期化できません');
        return;
    }
    console.log('マウスイベントリスナー登録完了');
}
svg.addEventListener("mousedown", (e) => {
    // ノードをクリックしていない場合のみパンを開始
    if (e.target === svg && !editingNode) {
        isPanning = true;
        panStartX = e.clientX - panX * zoom;
        panStartY = e.clientY - panY * zoom;
    }
});
svg.addEventListener("mousemove", (e) => {
    if (draggingNode) {
        const moveDistance = Math.abs(e.clientX - dragStartX) + Math.abs(e.clientY - dragStartY);
        // 移動距離が一定以上でドラッグ開始
        if (!isDragging && moveDistance > 5) {
            isDragging = true;
            console.log('🔄 ドラッグ開始:', draggingNode.text);
            // ドラッグ中のスタイルを適用
            const draggingElements = document.querySelectorAll('g');
            draggingElements.forEach(g => {
                var _a;
                if (((_a = g.querySelector('text')) === null || _a === void 0 ? void 0 : _a.textContent) === (draggingNode === null || draggingNode === void 0 ? void 0 : draggingNode.text)) {
                    g.classList.add('dragging');
                }
            });
        }
        if (isDragging) {
            // ノードドラッグ中（慣性アニメーション付き）
            const newX = (e.clientX - dragOffsetX) / zoom - panX;
            const newY = (e.clientY - dragOffsetY) / zoom - panY;
            
            // 目標位置を直接設定（ドラッグ中は即座移動）
            const deltaX = newX - draggingNode.x;
            const deltaY = newY - draggingNode.y;
            
            // メインノードと配下ノードを一緒に移動
            moveNodeWithChildren(draggingNode, deltaX, deltaY);
            
            // ドラッグ中は目標位置も更新
            function updateTargets(node, deltaX, deltaY) {
                initializeNodeAnimation(node);
                node.targetX = node.x;
                node.targetY = node.y;
                if (node.children) {
                    node.children.forEach(child => updateTargets(child, deltaX, deltaY));
                }
            }
            updateTargets(draggingNode, deltaX, deltaY);
            
            draw();
        }
    }
    else if (isPanning) {
        // キャンバスパン中
        panX = (e.clientX - panStartX) / zoom;
        panY = (e.clientY - panStartY) / zoom;
        draw();
    }
});
svg.addEventListener("mouseup", (e) => {
    if (draggingNode) {
        const clickDuration = Date.now() - clickStartTime;
        if (isDragging) {
            // ドラッグ終了時に重なっているノードを検索
            const overlappingNode = findOverlappingNode(draggingNode);
            
            if (overlappingNode) {
                // 重なったノードがある場合、階層構造を変更
                console.log('🔄 ドロップ検出:', draggingNode.text, 'を', overlappingNode.text, 'の配下に移動');
                
                // 元の親から削除
                removeNodeFromTree(draggingNode);
                
                // 新しい親の子として追加
                addNodeAsChild(overlappingNode, draggingNode);
                
                // レイアウトを再計算して描画（慣性アニメーション付き）
                layout(root);
                
                // 階層構造変更後にY座標順で整理
                setTimeout(() => reorganizeNodesByYPosition(), 200);
                
                console.log('✅ 階層構造変更完了');
            } else {
                // ドラッグ終了後、慣性で元の位置に戻る動きを追加
                function addInertia(node) {
                    initializeNodeAnimation(node);
                    // 目標位置は現在位置のままで最小限の調整
                    node.targetX = node.x;
                    node.targetY = node.y;
                    if (node.children) {
                        node.children.forEach(child => addInertia(child));
                    }
                }
                addInertia(draggingNode);
                
                // 衝突解消を慣性アニメーションで実行
                resolveAllCollisions();
                startAnimation();
                
                // ドラッグ終了後にY座標順で整理
                setTimeout(() => reorganizeNodesByYPosition(), 200);
            }
            
            // ドラッグ終了
            console.log('✅ ドラッグ終了:', draggingNode.text, 'を新しい位置に移動');
            pushHistory();
            // ドラッグスタイルをすべてクリア
            const draggingElements = document.querySelectorAll('g.dragging');
            draggingElements.forEach(g => g.classList.remove('dragging'));
        }
        else if (clickDuration < 300) {
            // 短時間のクリック：選択のみ
            console.log('📌 クリック選択:', draggingNode.text);
            setSelected(draggingNode);
        }
        draggingNode = null;
        isDragging = false;
        draw();
    }
    isPanning = false;
});
// 背景クリックで編集終了
svg.addEventListener("click", (e) => {
    // ノード以外をクリックした場合、編集を終了
    if (e.target === svg && editingNode) {
        finishEdit();
    }
});
// マウスホイールでズーム（画面フィットと同様の動作）
svg.addEventListener("wheel", (e) => {
    e.preventDefault();
    const rect = svg.getBoundingClientRect();
    const mouseX = e.clientX - rect.left;
    const mouseY = e.clientY - rect.top;
    // ズーム前のマウス位置に対応するワールド座標
    const worldX = (mouseX - panX * zoom) / zoom;
    const worldY = (mouseY - panY * zoom) / zoom;
    // ズーム倍率を調整（より滑らかに）
    const zoomFactor = e.deltaY > 0 ? 0.85 : 1.15;
    const newZoom = Math.max(0.1, Math.min(3.0, zoom * zoomFactor));
    // 画面中心を基準にズーム（フィット動作と統一）
    const centerX = rect.width / 2;
    const centerY = rect.height / 2;
    // ズーム後の調整
    panX = centerX / newZoom - worldX;
    panY = centerY / newZoom - worldY;
    zoom = newZoom;
    draw();
});
// ------------------------------
// ズーム・フィット機能
// ------------------------------
function zoomIn() {
    zoom = Math.min(maxZoom, zoom * 1.2);
    draw();
}
function zoomOut() {
    zoom = Math.max(minZoom, zoom / 1.2);
    draw();
}
function resetZoom() {
    zoom = 1.0;
    panX = 0;
    panY = 0;
    draw();
}
// ------------------------------
// キーボード操作
// ------------------------------
document.addEventListener("keydown", (e) => {
    var _a;
    // 編集中はキーボードショートカットを無効化
    if (editingNode) {
        e.stopPropagation();
        return;
    }
    // 文字キーで編集モード開始（スペースキーは除外）
    if (e.key.length === 1 && e.key !== ' ' && !e.ctrlKey && !e.altKey && !e.metaKey) {
        e.preventDefault();
        startInlineEditWithKey(selected, e.key);
        return;
    }
    if (e.ctrlKey && e.key === "s") {
        e.preventDefault();
        saveJSON();
        return;
    }
    if (e.ctrlKey && e.key === "o") {
        e.preventDefault();
        loadJSON();
        return;
    }
    if (e.ctrlKey && e.key === "z") {
        e.preventDefault();
        undo();
        return;
    }
    if (e.ctrlKey && e.key === "y") {
        e.preventDefault();
        redo();
        return;
    }
    // ズーム操作
    if (e.key === "+" || e.key === "=") {
        e.preventDefault();
        zoomIn();
        return;
    }
    if (e.key === "-") {
        e.preventDefault();
        zoomOut();
        return;
    }
    if (e.key === "0") {
        e.preventDefault();
        resetZoom();
        return;
    }
    if (e.key === "f" || e.key === "F") {
        e.preventDefault();
        fitToScreen();
        return;
    }
    // Ctrl + 矢印キーでノード順序変更
    if (e.ctrlKey && e.key === "ArrowUp") {
        e.preventDefault();
        moveNodeUp(selected);
        return;
    }
    if (e.ctrlKey && e.key === "ArrowDown") {
        e.preventDefault();
        moveNodeDown(selected);
        return;
    }
    // Shift + 矢印キーで従来の階層ナビゲーション
    if (e.shiftKey && e.key === "ArrowLeft") {
        e.preventDefault();
        const currentSelected = getSelected();
        if (currentSelected && currentSelected.parent) {
            setSelected(currentSelected.parent);
        }
        draw();
        return;
    }
    if (e.shiftKey && e.key === "ArrowRight") {
        e.preventDefault();
        const currentSelected = getSelected();
        if (currentSelected && currentSelected.children && currentSelected.children.length > 0) {
            setSelected(currentSelected.children[0]);
        }
        draw();
        return;
    }
    if (e.shiftKey && e.key === "ArrowUp") {
        e.preventDefault();
        const currentSelected = getSelected();
        if (currentSelected && currentSelected.parent) {
            const siblings = currentSelected.parent.children;
            const index = siblings.indexOf(currentSelected);
            if (index > 0) {
                setSelected(siblings[index - 1]);
            }
        }
        draw();
        return;
    }
    if (e.shiftKey && e.key === "ArrowDown") {
        e.preventDefault();
        const currentSelected = getSelected();
        if (currentSelected && currentSelected.parent) {
            const siblings = currentSelected.parent.children;
            const index = siblings.indexOf(currentSelected);
            if (index < siblings.length - 1) {
                setSelected(siblings[index + 1]);
            }
        }
        draw();
        return;
    }
    switch (e.key) {
        case "Enter":
            const currentSelected = getSelected();
            const newNode = addNode((_a = currentSelected.parent) !== null && _a !== void 0 ? _a : currentSelected);
            setSelected(newNode);
            break;
        case "Tab":
            e.preventDefault();
            // 選択したノードの子ノードを作成
            const currentSelectedForTab = getSelected();
            const childNode = addNode(currentSelectedForTab);
            setSelected(childNode);
            break;
        case "Delete":
        case "Backspace":
            console.log('⌨️ Deleteキーが押されました');
            const nodeToDelete = getSelected();
            console.log('🎯 削除対象ノード取得:', nodeToDelete ? nodeToDelete.text : 'null');
            deleteNode(nodeToDelete);
            break;
        case " ": // スペースキー
            e.preventDefault();
            const currentSelectedForCollapse = getSelected();
            if (currentSelectedForCollapse && currentSelectedForCollapse.children && currentSelectedForCollapse.children.length > 0) {
                const wasCollapsed = currentSelectedForCollapse.collapsed;
                
                // 折りたたみ状態を切り替え
                currentSelectedForCollapse.collapsed = !currentSelectedForCollapse.collapsed;
                console.log('📁 ノード折りたたみ切り替え:', currentSelectedForCollapse.text, 'collapsed:', currentSelectedForCollapse.collapsed);
                
                // 折りたたみ状態に応じてメッセージを表示
                if (currentSelectedForCollapse.collapsed) {
                    console.log('✅ ノードが折りたたまれました - 子ノードを非表示');
                } else {
                    console.log('✅ ノードが展開されました - 子ノードを表示');
                }
                
                pushHistory();
                saveCurrentMindMap();
                
                // 折りたたみ状態の変化に応じてレイアウトを調整
                setTimeout(() => {
                    if (wasCollapsed && !currentSelectedForCollapse.collapsed) {
                        // 展開時: 子ノードを再表示し、周辺ノードを調整
                        console.log('🔄 展開時のレイアウト調整を開始');
                        adjustLayoutAfterExpansion(currentSelectedForCollapse);
                    } else if (!wasCollapsed && currentSelectedForCollapse.collapsed) {
                        // 折りたたみ時: 周辺ノードを詰める
                        console.log('🔄 折りたたみ時のレイアウト調整を開始');
                        adjustLayoutAfterCollapse(currentSelectedForCollapse);
                    }
                }, 50);
                
                draw();
            } else {
                console.log('⚠️ 折りたたみ対象の子ノードがありません:', currentSelectedForCollapse ? currentSelectedForCollapse.text : 'null');
            }
            break;
        case "ArrowUp":
            // 視覚的に上にあるノードに移動
            const currentSelectedUp = getSelected();
            const upNode = findClosestNodeInDirection(currentSelectedUp, 'up');
            if (upNode)
                setSelected(upNode);
            break;
        case "ArrowDown":
            // 視覚的に下にあるノードに移動
            const currentSelectedDown = getSelected();
            const downNode = findClosestNodeInDirection(currentSelectedDown, 'down');
            if (downNode)
                setSelected(downNode);
            break;
        case "ArrowLeft":
            // 視覚的に左にあるノードに移動
            const currentSelectedLeft = getSelected();
            const leftNode = findClosestNodeInDirection(currentSelectedLeft, 'left');
            if (leftNode)
                setSelected(leftNode);
            break;
        case "ArrowRight":
            // 視覚的に右にあるノードに移動
            const currentSelectedRight = getSelected();
            const rightNode = findClosestNodeInDirection(currentSelectedRight, 'right');
            if (rightNode)
                setSelected(rightNode);
            break;
        case "F2":
            e.preventDefault();
            if (!editingNode) {
                const currentSelectedEdit = getSelected();
                startInlineEdit(currentSelectedEdit, null);
            }
            break;
        case "+":
        case "=":
            // ズームイン
            e.preventDefault();
            zoom = Math.min(zoom * 1.2, 3);
            draw();
            return;
        case "-":
            // ズームアウト  
            e.preventDefault();
            zoom = Math.max(zoom / 1.2, 0.1);
            draw();
            return;
        case "0":
            // ズームリセット
            e.preventDefault();
            zoom = 1;
            panX = 0;
            panY = 0;
            draw();
            return;
        case "f":
        case "F":
            // 画面に合わせる
            e.preventDefault();
            fitToScreen();
            return;
    }
    draw();
});
// ノードを画面中央に移動
function centerNodeOnScreen(node) {
    // SVGのサイズを取得
    const rect = svg.getBoundingClientRect();
    const screenCenterX = rect.width / 2;
    const screenCenterY = rect.height / 2;
    // ノードを画面中央に配置するためのパン値を計算
    panX = (screenCenterX / zoom) - node.x;
    panY = (screenCenterY / zoom) - node.y;
}
// 画面に合わせる：全ノードが見えるようにズームとパンを調整
function fitToScreen() {
    const allNodes = getAllNodes(root);
    if (allNodes.length === 0)
        return;
    // 全ノードの境界ボックスを計算
    let minX = allNodes[0].x;
    let maxX = allNodes[0].x;
    let minY = allNodes[0].y;
    let maxY = allNodes[0].y;
    allNodes.forEach(node => {
        const nodeWidth = Math.max(120, node.text.length * 8) / 2; // 動的幅の半分
        minX = Math.min(minX, node.x - nodeWidth);
        maxX = Math.max(maxX, node.x + nodeWidth);
        minY = Math.min(minY, node.y - 20); // ノード高さの半分
        maxY = Math.max(maxY, node.y + 20);
    });
    // マージンを追加
    const margin = 50;
    minX -= margin;
    maxX += margin;
    minY -= margin;
    maxY += margin;
    const contentWidth = maxX - minX;
    const contentHeight = maxY - minY;
    // SVGのサイズを取得（画面サイズ）
    const rect = svg.getBoundingClientRect();
    const screenWidth = rect.width;
    const screenHeight = rect.height;
    // 画面に収まるズームレベルを計算
    const zoomX = screenWidth / contentWidth;
    const zoomY = screenHeight / contentHeight;
    zoom = Math.min(zoomX, zoomY, 1); // 最大1倍まで
    // コンテンツの中心が画面中央に来るようにパンを調整
    const centerX = (minX + maxX) / 2;
    const centerY = (minY + maxY) / 2;
    panX = (screenWidth / zoom / 2) - centerX;
    panY = (screenHeight / zoom / 2) - centerY;
    draw();
}
// ------------------------------
// アプリケーション初期化
// ------------------------------
// 初期化関数
function initializeApplication() {
    console.log('アプリケーション初期化開始');
    // DOM要素の確認
    if (!svg) {
        console.error('SVG要素が見つかりません');
        return;
    }
    // マインドマップデータを読み込み
    loadMindMapsFromStorage();
    // UI初期化
    initializeMindMapUI();
    initializeLayoutUI();
    initializeMouseEvents(); // マウスイベントを明示的に初期化
    setupGlobalMenuListeners();
    updateUI();
    // 初期履歴（マインドマップ管理に統合済み）
    if (undoHistory.length === 0) {
        pushHistory();
    }
    draw();
    console.log('アプリケーション初期化完了');
    // 操作方法を表示
    setTimeout(() => {
        console.log('');
        console.log('🎯 ========== マインドマップ操作方法 ==========');
        console.log('');
        console.log('📱 マウス操作:');
        console.log('  • クリック: ノードを選択');
        console.log('  • Ctrl+クリック: 選択ノードを画面中央に移動');
        console.log('  • ダブルクリック: テキスト編集');
        console.log('  • ドラッグ: ノード移動');
        console.log('  • 右クリック: パン操作');
        console.log('  • ホイール: ズーム操作');
        console.log('');
        console.log('⌨️  キーボード操作:');
        console.log('  • 矢印キー: 視覚的にノード移動');
        console.log('  • Shift+矢印キー: 階層ナビゲーション');
        console.log('  • Enter: 兄弟ノード作成');
        console.log('  • Tab: 子ノード作成');
        console.log('  • Delete/Backspace: ノード削除');
        console.log('  • F2: テキスト編集');
        console.log('');
        console.log('💡 レイアウト切り替え:');
        console.log('  • 左上ハンバーガーメニューから選択');
        console.log('  • Radial（放射状）/ Left-Right（左右分岐）/ Tree（片側ツリー）');
        console.log('');
        console.log('🔧 デバッグ機能:');
        console.log('  • debugMenu.toggle(): メニュー開閉テスト');
        console.log('  • debugMenu.testLayoutButtons(): レイアウトボタンテスト');
        console.log('  • debugMenu.switchToRadial/LeftRight/Tree(): 手動レイアウト切替');
        console.log('');
        console.log('============================================');
        console.log('');
    }, 1000);
}
// DOM読み込み完了後に初期化実行
// 複数のタイミングで確実に実行
console.log('script開始、DOM状態:', document.readyState);
function ensureInitialization() {
    console.log('確実な初期化実行開始');
    // 複数回実行防止
    if (window.mindmapInitialized) {
        console.log('既に初期化済み、スキップ');
        return;
    }
    // DOM要素が存在するか再確認
    const svg = document.getElementById('mindmap');
    const slideMenu = document.getElementById('slide-menu');
    const menuToggle = document.getElementById('menu-toggle');
    console.log('重要要素チェック:', {
        svg: !!svg,
        slideMenu: !!slideMenu,
        menuToggle: !!menuToggle
    });
    if (!svg || !slideMenu || !menuToggle) {
        console.log('まだDOM要素が不完全、1秒後に再試行');
        setTimeout(ensureInitialization, 1000);
        return;
    }
    console.log('DOM要素揃った、初期化実行');
    initializeApplication();
    window.mindmapInitialized = true;
}
if (document.readyState === 'loading') {
    console.log('DOMまだ読み込み中、DOMContentLoadedで待機');
    document.addEventListener('DOMContentLoaded', ensureInitialization);
}
else {
    console.log('DOM既に読み込み済み、即座に初期化チェック');
    setTimeout(ensureInitialization, 100);
}
// フォールバックとしてwindow.onloadも設定
window.addEventListener('load', () => {
    console.log('window.load発火、念のため確実な初期化実行');
    ensureInitialization();
});
// ------------------------------
// UI操作（スライドメニュー）
// ------------------------------
// レイアウトUI初期化
function initializeLayoutUI() {
    console.log('レイアウトUI初期化開始');
    // DOM要素の存在確認（詳細）
    console.log('HTML全体:', document.documentElement.innerHTML.length > 0 ? 'OK' : 'NG');
    console.log('body要素:', !!document.body);
    const slideMenu = document.getElementById('slide-menu');
    const menuToggle = document.getElementById('menu-toggle');
    const menuClose = document.getElementById('menu-close');
    const radialBtn = document.getElementById('radial-layout');
    const leftrightBtn = document.getElementById('leftright-layout');
    const treeBtn = document.getElementById('tree-layout');
    console.log('DOM要素取得結果:', {
        slideMenu: !!slideMenu,
        menuToggle: !!menuToggle,
        menuClose: !!menuClose,
        radialBtn: !!radialBtn,
        leftrightBtn: !!leftrightBtn,
        treeBtn: !!treeBtn
    });
    if (!slideMenu) {
        console.error('slide-menu要素が見つかりません');
        return;
    }
    if (!menuToggle) {
        console.error('menu-toggle要素が見つかりません');
        return;
    }
    if (!menuClose) {
        console.error('menu-close要素が見つかりません');
        return;
    }
    // 初期状態を強制設定
    console.log('初期状態設定開始');
    slideMenu.classList.remove('open'); // 確実に開いていない状態にする
    menuToggle.classList.remove('menu-open');
    // CSS適用状態をチェック
    const slideMenuStyle = window.getComputedStyle(slideMenu);
    console.log('slideMenuのCSS状態:', {
        position: slideMenuStyle.position,
        left: slideMenuStyle.left,
        width: slideMenuStyle.width,
        height: slideMenuStyle.height,
        zIndex: slideMenuStyle.zIndex,
        display: slideMenuStyle.display,
        visibility: slideMenuStyle.visibility
    });
    const toggleStyle = window.getComputedStyle(menuToggle);
    console.log('menuToggleのCSS状態:', {
        position: toggleStyle.position,
        left: toggleStyle.left,
        top: toggleStyle.top,
        width: toggleStyle.width,
        height: toggleStyle.height,
        zIndex: toggleStyle.zIndex,
        display: toggleStyle.display,
        visibility: toggleStyle.visibility,
        pointerEvents: toggleStyle.pointerEvents
    });
    console.log('初期状態設定完了');
    // トグルボタンの確実なイベント登録
    console.log('メニュートグルボタンにシンプルなリスナー登録開始');
    console.log('ボタン要素情報:', {
        id: menuToggle.id,
        className: menuToggle.className,
        tagName: menuToggle.tagName,
        offsetLeft: menuToggle.offsetLeft,
        offsetTop: menuToggle.offsetTop,
        offsetWidth: menuToggle.offsetWidth,
        offsetHeight: menuToggle.offsetHeight
    });
    // シンプルなクリックイベント
    menuToggle.onclick = function (e) {
        console.log('==================');
        console.log('onclick イベント発火!');
        console.log('現在のclassName:', slideMenu.className);
        e === null || e === void 0 ? void 0 : e.preventDefault();
        e === null || e === void 0 ? void 0 : e.stopPropagation();
        if (slideMenu.classList.contains('open')) {
            slideMenu.classList.remove('open');
            menuToggle.classList.remove('menu-open');
            console.log('メニューを閉じました');
        }
        else {
            slideMenu.classList.add('open');
            menuToggle.classList.add('menu-open');
            console.log('メニューを開きました');
        }
        console.log('変更後のclassName:', slideMenu.className);
        console.log('==================');
        return false;
    };
    console.log('onclick イベント登録完了');
    // テスト用：手動でスライドメニューを開く
    console.log('3秒後にスライドメニューを自動で開きます（CSS動作確認）');
    setTimeout(() => {
        console.log('スライドメニューを手動で開きます');
        slideMenu.classList.add('open');
        console.log('手動オープン後のclassName:', slideMenu.className);
        // さらに3秒後に閉じる
        setTimeout(() => {
            console.log('スライドメニューを手動で閉じます');
            slideMenu.classList.remove('open');
            console.log('手動クローズ後のclassName:', slideMenu.className);
        }, 3000);
    }, 3000);
    console.log('メニュークローズボタンにリスナー登録');
    menuClose.addEventListener('click', (e) => {
        console.log('メニュークローズクリック');
        e.preventDefault();
        slideMenu.classList.remove('open');
        menuToggle.classList.remove('menu-open');
    });
    if (!radialBtn || !leftrightBtn || !treeBtn) {
        console.error('レイアウトボタンが見つかりません');
        // スライドメニュートグルは動作するが、レイアウトボタンは無効
        return;
    }
    // レイアウトボタンのstyle情報をチェック
    [radialBtn, leftrightBtn, treeBtn].forEach((btn, index) => {
        const btnName = ['Radial', 'LeftRight', 'Tree'][index];
        const style = window.getComputedStyle(btn);
        console.log(`${btnName}ボタンのCSS状態:`, {
            display: style.display,
            visibility: style.visibility,
            pointerEvents: style.pointerEvents,
            zIndex: style.zIndex,
            position: style.position,
            cursor: style.cursor,
            width: style.width,
            height: style.height
        });
    });
    console.log('レイアウトボタンにシンプルなイベントリスナー登録');
    // シンプルなonclickイベントで登録
    radialBtn.onclick = function (e) {
        console.log('==================');
        console.log('Radialレイアウトボタンクリック!');
        e === null || e === void 0 ? void 0 : e.preventDefault();
        e === null || e === void 0 ? void 0 : e.stopPropagation();
        switchLayout('radial');
        console.log('Radialレイアウト適用完了');
        console.log('==================');
        return false;
    };
    leftrightBtn.onclick = function (e) {
        console.log('==================');
        console.log('Left-Rightレイアウトボタンクリック!');
        e === null || e === void 0 ? void 0 : e.preventDefault();
        e === null || e === void 0 ? void 0 : e.stopPropagation();
        switchLayout('leftright');
        console.log('Left-Rightレイアウト適用完了');
        console.log('==================');
        return false;
    };
    treeBtn.onclick = function (e) {
        console.log('==================');
        console.log('Treeレイアウトボタンクリック!');
        e === null || e === void 0 ? void 0 : e.preventDefault();
        e === null || e === void 0 ? void 0 : e.stopPropagation();
        switchLayout('tree');
        console.log('Treeレイアウト適用完了');
        console.log('==================');
        return false;
    };
    console.log('レイアウトUI初期化完了');
    // グローバルデバッグ関数を追加
    window.debugMenu = {
        toggle: () => {
            console.log('手動トグル実行');
            slideMenu.classList.toggle('open');
            console.log('手動トグル後のクラス:', slideMenu.className);
        },
        open: () => {
            console.log('手動オープン実行');
            slideMenu.classList.add('open');
            console.log('手動オープン後のクラス:', slideMenu.className);
        },
        close: () => {
            console.log('手動クローズ実行');
            slideMenu.classList.remove('open');
            console.log('手動クローズ後のクラス:', slideMenu.className);
        },
        checkButton: () => {
            const btn = document.getElementById('menu-toggle');
            console.log('ボタン状態確認:', {
                存在: !!btn,
                表示: btn ? getComputedStyle(btn).display : 'N/A',
                位置: btn ? `${btn.offsetLeft}, ${btn.offsetTop}` : 'N/A',
                サイズ: btn ? `${btn.offsetWidth}x${btn.offsetHeight}` : 'N/A',
                zIndex: btn ? getComputedStyle(btn).zIndex : 'N/A',
                pointerEvents: btn ? getComputedStyle(btn).pointerEvents : 'N/A'
            });
        },
        testLayoutButtons: () => {
            console.log('レイアウトボタンテスト開始');
            const buttons = [
                { id: 'radial-layout', name: 'Radial' },
                { id: 'leftright-layout', name: 'LeftRight' },
                { id: 'tree-layout', name: 'Tree' }
            ];
            buttons.forEach(({ id, name }) => {
                const btn = document.getElementById(id);
                console.log(`${name}ボタン状態:`, {
                    存在: !!btn,
                    クリック可能: btn ? btn.onclick !== null : false,
                    表示: btn ? getComputedStyle(btn).display : 'N/A',
                    ポインタ: btn ? getComputedStyle(btn).pointerEvents : 'N/A',
                    カーソル: btn ? getComputedStyle(btn).cursor : 'N/A'
                });
                if (btn) {
                    console.log(`${name}ボタンを手動クリックテスト`);
                    btn.click();
                }
            });
        },
        switchToRadial: () => {
            console.log('手動でRadialに切り替え');
            switchLayout('radial');
        },
        switchToLeftRight: () => {
            console.log('手動でLeft-Rightに切り替え');
            switchLayout('leftright');
        },
        switchToTree: () => {
            console.log('手動でTreeに切り替え');
            switchLayout('tree');
        }
    };
    console.log('グローバルデバッグ関数を window.debugMenu に追加しました');
    console.log('使用方法:');
    console.log('  debugMenu.toggle() - メニー開閉');
    console.log('  debugMenu.open() - メニーを開く');
    console.log('  debugMenu.close() - メニーを閉じる');
    console.log('  debugMenu.checkButton() - ハンバーガーボタン状態確認');
    console.log('  debugMenu.testLayoutButtons() - レイアウトボタン全チェック');
    console.log('  debugMenu.switchToRadial() - 手動でRadial切替');
    console.log('  debugMenu.switchToLeftRight() - 手動でLeftRight切替');
    console.log('  debugMenu.switchToTree() - 手動でTree切替');
}
// レイアウト切り替え
function switchLayout(newLayout) {
    // アニメーションを停止してから切り替え
    stopAnimation();
    
    currentLayout = newLayout;
    // アクティブボタンの更新
    updateLayoutButtons();
    // レイアウトを再計算して描画（慣性アニメーション付き）
    layout(root);
    closeMenu();
}
// レイアウト切り替え（メニューを閉じない版）
function switchLayoutSilent(newLayout) {
    // アニメーションを停止してから切り替え
    stopAnimation();
    
    currentLayout = newLayout;
    // アクティブボタンの更新
    updateLayoutButtons();
    // レイアウトを再計算して描画（慣性アニメーション付き）
    layout(root);
}
// レイアウトボタンの状態更新
function updateLayoutButtons() {
    document.querySelectorAll('.layout-btn').forEach(btn => btn.classList.remove('active'));
    const radialBtn = document.getElementById('radial-layout');
    const leftrightBtn = document.getElementById('leftright-layout');
    const treeBtn = document.getElementById('tree-layout');
    switch (currentLayout) {
        case 'radial':
            radialBtn === null || radialBtn === void 0 ? void 0 : radialBtn.classList.add('active');
            break;
        case 'leftright':
            leftrightBtn === null || leftrightBtn === void 0 ? void 0 : leftrightBtn.classList.add('active');
            break;
        case 'tree':
            treeBtn === null || treeBtn === void 0 ? void 0 : treeBtn.classList.add('active');
            break;
    }
}
// switchLayout関数でメニュー閉じる処理を修正
function closeMenu() {
    const slideMenu = document.getElementById('slide-menu');
    const menuToggle = document.getElementById('menu-toggle');
    if (slideMenu && menuToggle) {
        slideMenu.classList.remove('open');
        menuToggle.classList.remove('menu-open');
    }
}
// ------------------------------
// マインドマップ管理機能
// ------------------------------
// 新規マインドマップ作成
function createNewMindMap(name) {
    const id = 'mindmap_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    const timestamp = Date.now();
    const newMindMap = {
        id,
        name: name || `マインドマップ ${mindMaps.length + 1}`,
        rootNode: {
            id: "root",
            text: "ルートアイテム",
            children: [],
            parent: null,
            x: 800,
            y: 450
        },
        layout: 'radial',
        createdAt: timestamp,
        updatedAt: timestamp
    };
    mindMaps.push(newMindMap);
    saveMindMapsToStorage();
    return newMindMap;
}
// マインドマップを切り替え
function switchToMindMap(id) {
    console.log(`switchToMindMap開始: ${id}`);
    const mindMap = mindMaps.find(m => m.id === id);
    if (!mindMap) {
        console.error(`マインドマップが見つかりません: ${id}`);
        return;
    }
    console.log(`切り替え先: ${mindMap.name}`);
    // 現在のマインドマップを保存
    saveCurrentMindMap();
    // 新しいマインドマップを読み込み
    activeMindMapId = id;
    root = mindMap.rootNode;
    currentLayout = mindMap.layout;
    setSelected(root); // 安全な設定を使用
    // 履歴をクリア
    undoHistory = [];
    historyIndex = -1;
    pushHistory();
    // レイアウトボタンの状態を更新
    switchLayoutSilent(currentLayout);
    // レイアウトを適用
    layout(root);
    draw();
    updateUI();
    console.log(`switchToMindMap完了: ${mindMap.name}`);
}
// 現在のマインドマップを保存
function saveCurrentMindMap() {
    if (!activeMindMapId)
        return;
    const mindMap = mindMaps.find(m => m.id === activeMindMapId);
    if (mindMap) {
        try {
            // 循環参照を防ぐためreplacer関数を使用
            mindMap.rootNode = JSON.parse(JSON.stringify(root, replacer));
            mindMap.layout = currentLayout;
            mindMap.updatedAt = Date.now();
            saveMindMapsToStorage();
            console.log('マインドマップ保存完了:', mindMap.name);
        }
        catch (error) {
            console.error('マインドマップ保存エラー:', error);
        }
    }
}
// マインドマップを削除
function deleteMindMap(id) {
    if (mindMaps.length <= 1) {
        alert('最後のマインドマップは削除できません');
        return;
    }
    const index = mindMaps.findIndex(m => m.id === id);
    if (index === -1)
        return;
    mindMaps.splice(index, 1);
    // 削除したマインドマップがアクティブだった場合
    if (id === activeMindMapId) {
        const newActiveMindMap = mindMaps[0];
        switchToMindMap(newActiveMindMap.id);
    }
    else {
        saveMindMapsToStorage();
        updateUI();
    }
}
// ローカルストレージに保存
function saveMindMapsToStorage() {
    try {
        const data = {
            mindMaps,
            activeMindMapId
        };
        // 循環参照を防ぐためreplacer関数を使用
        localStorage.setItem('mindMaps', JSON.stringify(data, replacer));
        console.log('マインドマップデータ保存完了');
    }
    catch (error) {
        console.error('マインドマップデータの保存に失敗:', error);
    }
}
// ローカルストレージから読み込み
function loadMindMapsFromStorage() {
    const saved = localStorage.getItem('mindMaps');
    if (!saved) {
        // 初回起動時：デフォルトマインドマップを作成
        console.log('初回起動、デフォルトマインドマップ作成');
        const defaultMindMap = createNewMindMap('マインドマップ 1');
        activeMindMapId = defaultMindMap.id;
        root = defaultMindMap.rootNode;
        return;
    }
    try {
        console.log('保存済みデータを読み込み中');
        const data = JSON.parse(saved);
        mindMaps = data.mindMaps || [];
        activeMindMapId = data.activeMindMapId || '';
        
        // parent参照を復元（安全なコピーで元データを保護）
        mindMaps.forEach((mindMap, index) => {
            console.log(`マインドマップ${index + 1}のnode参照を復元中: ${mindMap.name}`);
            if (mindMap.rootNode) {
                mindMap.rootNode = restoreParentReferences(JSON.parse(JSON.stringify(mindMap.rootNode)));
            }
        });
        
        // アクティブマインドマップを設定
        if (activeMindMapId && mindMaps.find(m => m.id === activeMindMapId)) {
            console.log(`アクティブマインドマップに切り替え: ${activeMindMapId}`);
            const activeMindMap = mindMaps.find(m => m.id === activeMindMapId);
            root = activeMindMap.rootNode;
            currentLayout = activeMindMap.layout || 'radial';
        } else if (mindMaps.length > 0) {
            console.log('最初のマインドマップをアクティブに設定');
            activeMindMapId = mindMaps[0].id;
            root = mindMaps[0].rootNode;
            currentLayout = mindMaps[0].layout || 'radial';
        } else {
            // データが破損している場合：新規作成
            console.log('データ破損、デフォルトマインドマップ作成');
            const defaultMindMap = createNewMindMap('マインドマップ 1');
            activeMindMapId = defaultMindMap.id;
            root = defaultMindMap.rootNode;
        }
    } catch (e) {
        console.error('マインドマップデータの読み込みに失敗:', e);
        // エラー時：新規作成
        const defaultMindMap = createNewMindMap('マインドマップ 1');
        activeMindMapId = defaultMindMap.id;
        root = defaultMindMap.rootNode;
    }
}

// parent参照を安全に復元（元データを変更しない）
function restoreParentReferences(node, parent = null) {
    node.parent = parent;
    if (node.children && Array.isArray(node.children)) {
        node.children.forEach(child => {
            restoreParentReferences(child, node);
        });
    }
    return node;
}
// UI更新
function updateUI() {
    updateMindMapSelector();
    updateMindMapList();
}
// マインドマップセレクターを更新（削除されたため無効化）
function updateMindMapSelector() {
    // HTML要素が削除されたため、この関数は無効化
    console.log('updateMindMapSelector: 無効化済み');
    return;
}
// マインドマップリストを更新（削除されたため無効化）
function updateMindMapList() {
    // HTML要素が削除されたため、この関数は無効化
    console.log('updateMindMapList: 無効化済み');
    return;
}
// マインドマップリストの表示/非表示（無効化）
function toggleMindMapList() {
    console.log('toggleMindMapList: 無効化済み');
    return;
}
function showMindMapList() {
    console.log('showMindMapList: 無効化済み');
    return;
}
function hideMindMapList() {
    console.log('hideMindMapList: 無効化済み');
    return;
}
// ------------------------------
// マインドマップUI初期化
// ------------------------------
// マインドマップ管理UIの要素取得と初期化（無効化）
function initializeMindMapUI() {
    console.log('マインドマップUI初期化: 無効化済み');
    return;
}
// グローバルイベントリスナー設定
function setupGlobalMenuListeners() {
    document.addEventListener('click', (e) => {
        const target = e.target;
        const slideMenu = document.getElementById('slide-menu');
        const menuToggle = document.getElementById('menu-toggle');
        
        // スライドメニューを閉じる
        if (slideMenu && menuToggle &&
            !slideMenu.contains(target) &&
            !menuToggle.contains(target)) {
            if (slideMenu.classList.contains('open')) {
                closeMenu();
            }
        }
    });
}
// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM読み込み完了 - 初期化開始');
    
    // URLパラメータで指定されたマインドマップがあれば読み込み
    const loaded = loadSpecifiedMindMap();
    
    if (!loaded) {
        // 通常の初期化処理（統合版）
        loadMindMapsFromStorage();
    }
    
    // 初期慈性アニメーションの設定
    if (!root.x) {
        root.x = root.y = 0;
    }
    initializeNodeAnimation(root);
    
    // selected変数の初期化
    if (typeof selected === 'undefined' || !selected) {
        setSelected(root);
    } else {
        setSelected(selected);
    }
    
    // 初期履歴
    if (undoHistory.length === 0) {
        pushHistory();
    }
    
    // 初期描画
    if (typeof layout === 'function') layout(root);
    if (typeof draw === 'function') draw();
    
    console.log('初期化完了 - root:', root);
});

// 定期的に現在のマインドマップを保存
setInterval(() => {
    saveCurrentMindMap();
}, 5000); // 5秒毎に自動保存
