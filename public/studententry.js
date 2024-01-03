document.addEventListener('DOMContentLoaded', function () {
    const studentForm = document.getElementById('studentForm');
    const infoModal = document.getElementById('infoModal');
    const modalContent = document.getElementById('modalContent');

    // Extract bookId from the URL query parameters
    const urlSearchParams = new URLSearchParams(window.location.search);
    const bookId = urlSearchParams.get('bookId');

    // Set the bookId value in the hidden input field
    document.getElementById('bookId').value = bookId;

    studentForm.addEventListener('submit', function (event) {
        event.preventDefault();

        const name = document.getElementById('name').value;
        const phone = document.getElementById('phone').value;
        const department = document.getElementById('department').value;
        const academicYear = document.getElementById('academicYear').value;
        const paymentMethod = document.getElementById('Payment Method').value;

        // Update modal content including the book ID
        modalContent.innerHTML = `Student Information:
            <br>Name: ${name}
            <br>Phone Number: ${phone}
            <br>Department: ${department}
            <br>Academic Year: ${academicYear}
            <br>Book ID: ${bookId}
            <br>Payment Method: ${paymentMethod}`;

        // Display modal
        infoModal.style.display = 'block';

        // Close modal after 3 seconds (3000 milliseconds)
        setTimeout(closeModal, 3000);

        // Place order logic - you can extend this to send the order details to the server
        placeOrder({
            bookId: bookId,
            name: name,
            phone: phone,
            department: department,
            academicYear: academicYear,
            paymentMethod: paymentMethod
        });
    });

    // Close modal function
    window.closeModal = function () {
        infoModal.style.display = 'none';
    };

    // Function to place the order (you can customize this)
    function placeOrder(orderDetails) {
        // Send order details to the server (replace the URL with your actual server endpoint)
        fetch('/place-order', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(orderDetails)
        })
        .then(response => response.json())
        .then(data => {
            console.log('Order placed successfully:', data);
            // You can redirect or show a success message to the user
        })
        .catch(error => {
            console.error('Error placing order:', error);
            // Handle error scenario
        });
    }
});
