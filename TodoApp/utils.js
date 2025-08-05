const TODO_KEY = 'todos';

export const saveTodosInLocalStorage = (todos) => {
     localStorage.setItem(TODO_KEY, JSON.stringify(todos));
}
export const getTodosFromLocalStorage = () => {
     return JSON.parse(localStorage.getItem(TODO_KEY)) || [];
}
export const getDateRepresentation = (date) => {
     return Intl.DateTimeFormat('ru-RU', {
          day: 'numeric',
          month: 'numeric',
          year: 'numeric'
     }).format(date)
}