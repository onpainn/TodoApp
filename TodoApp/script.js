import {saveTodosInLocalStorage, getTodosFromLocalStorage, getDateRepresentation} from './utils.js'

const addTodoInput = document.querySelector("[data-add-todo-input]");
const addTodoBtn = document.querySelector("[data-add-todo-btn]");
const todoContainer = document.querySelector("[data-todo-container]");
const todoTemplate = document.querySelector("[data-todo-template]");
const searchTodo = document.querySelector("[data-search-todo-input]");

let todoList = getTodosFromLocalStorage();
let filteredTodosList = [];

addTodoBtn.addEventListener('click', () => {
    if(addTodoInput.value.trim()){
        const newTodo = {
            id: Date.now(),
            text: addTodoInput.value.trim(),
            isCompleted: false,
            createdAt: getDateRepresentation(new Date()),
        }
        todoList.push(newTodo);
        saveTodosInLocalStorage(todoList);
        addTodoInput.value = '';
        renderTodo()
    }
})


addTodoInput.addEventListener('input', (e) => {
    if(searchTodo.value.trim()){
        searchTodo.value = '';
        renderTodo()
    }
})







searchTodo.addEventListener('input', (e) => {
    const searchValue = e.target.value.trim();
    filterAndRenderTodo(searchValue);
})


const filterAndRenderTodo = (searchValue) => {
    filteredTodosList = todoList.filter((t) => {
        return t.text.includes(searchValue);
    })
    renderFilterTodo()
}

const createTodoLayout = (todo) => {
    const todoElement = document.importNode(todoTemplate.content, true);

    const checkbox = todoElement.querySelector('[data-todo-checkbox]');
    checkbox.checked = todo.isCompleted;
    const todoText = todoElement.querySelector('[data-todo-text]');
    todoText.textContent = todo.text;
    const todoDate = todoElement.querySelector('[data-todo-date]'); 
    todoDate.textContent = todo.createdAt;
    const removeTodoBtn = todoElement.querySelector('[data-remove-todo-btn]');
    removeTodoBtn.disabled = !todo.isCompleted;
    checkbox.addEventListener('change', (e) => {
        todoList = todoList.map((t) => {
            if (t.id === todo.id){
                t.isCompleted = e.target.checked;
            }
            return t;
        })
        saveTodosInLocalStorage(todoList);
        
        if(searchTodo.value.trim()){
            filteredTodosList = todoList.filter((t) => {
                return t.text.includes(searchTodo.value.trim());
            })
            filterAndRenderTodo(searchTodo.value.trim());
        }else{
            renderTodo()
        }
    })


    removeTodoBtn.addEventListener('click', () => {
         todoList = todoList.filter((t) => {
            if (t.id !== todo.id){
                return t;
            }
        })
        saveTodosInLocalStorage(todoList);
        if(searchTodo.value.trim()){
            filteredTodosList = todoList.filter((t) => {
                return t.text.includes(searchTodo.value.trim());
            })
            filterAndRenderTodo(searchTodo.value.trim());
        }else{
            renderTodo()
        }
    })

    return todoElement;
}
const renderFilterTodo = () => {
    todoContainer.innerHTML = '';

    if(filteredTodosList.length === 0){
        todoContainer.innerHTML = '<p class="noTodo">No todos found...</p>';
        return;
    }
    filteredTodosList.forEach((todo) => {
        const todoElement = createTodoLayout(todo);
        todoContainer.append(todoElement)
    });
}
const renderTodo = () => {
    todoContainer.innerHTML = '';

    if(todoList.length === 0){
        todoContainer.innerHTML = '<p class="noTodo">You have not a todos...</p>';
        return;
    }
    todoList.forEach((todo) => {
        const todoElement = createTodoLayout(todo);
        todoContainer.append(todoElement)
    });
}
renderTodo()